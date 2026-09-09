SET TRANSACTION ISOLATION LEVEL READ COMMITTED
GOGO
DELETE FROM TRGDISA
GOGO
/* Increasing suffix field size in OrdDsg table */
IF EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('OrdDsg') AND upper(COLUMN_NAME) = upper('OdSfx'))
	Alter table OrdDsg alter column OdSfx varchar(15) not null 
GOGO

/* Increasing suffix field size in CustDsg table */
IF EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('CustDsg') AND upper(COLUMN_NAME) = upper('CdOurSfx'))
	Alter table CustDsg alter column CdOurSfx varchar(15) not null 
GOGO

/* Increasing suffix field size in Bag table */
IF EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('Bag') AND upper(COLUMN_NAME) = upper('BOdSfx'))
	Alter table Bag alter column BOdSfx varchar(15) not null 
GOGO

/* Increasing suffix field size in Fgd table */
IF EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('Fgd') AND upper(COLUMN_NAME) = upper('FdSfx'))
	Alter table Fgd alter column FdSfx varchar(15) not null 
GOGO

/* Increasing suffix field size in InvDsg table */
IF EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('InvDsg') AND upper(COLUMN_NAME) = upper('IdSfx'))
	Alter table InvDsg alter column IdSfx varchar(15) not null 
GOGO

/* Increasing suffix field size in DmYyMm table */
IF EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('DmYyMm') AND upper(COLUMN_NAME) = upper('DOdSfx'))
	Alter table DmYyMm alter column DOdSfx varchar(15) not null 
GOGO

/* Increasing suffix field size in TcStr table */
Update TcStr set TsMaxLen =15 
	WHERE TsTbl ='RmRt' and TsTyp ='MRD' and TsSTyp ='' and TsIdName ='RrDmSfx' 
GOGO
Update TcStr set TsMaxLen =15 
	WHERE TsTbl ='RmRt' and TsTyp ='MRP' and TsSTyp ='' and TsIdName ='RrDmSfx' 
GOGO




/* Deleting index from RmRt table */
IF  EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[RmRt]') AND name = N'UK_RmRt')
DROP INDEX [UK_RmRt] ON [dbo].[RmRt] WITH ( ONLINE = OFF )
GOGO
/* Deleting RrAllKeys column from RmRt table */
IF EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('RmRt') AND upper(COLUMN_NAME) = upper('RrAllKeys'))
	ALTER TABLE RmRt DROP COLUMN RrAllKeys
GOGO

/* Increasing suffix field size in RmRt table */
IF EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('RmRt') AND upper(COLUMN_NAME) = upper('RrDmSfx'))
	Alter table RmRt alter column RrDmSfx varchar(15) not null 
GOGO

/* Adding RrAllKeys column to RmRt table */
IF not EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('RmRt') AND upper(COLUMN_NAME) = upper('RrAllKeys'))
ALTER TABLE RmRt ADD RrAllKeys AS ((((((((((((((((((((((((((((((((([RrTcTyp]+'\')+[RrCmCtg])+'\')+[RrCmCd])+'\')+[RrCmCurCd])+'\')+[RrCtg])+'\')+[RrSctg])+'\')+[RrSTWGrp])+'\')+[RrCd])+'\')+[RrRmKT])+'\')+[RrDmCtg])+'\')+[RrDmCd])+'\')+[RrDmSfx])+'\')+[RrDmSz])+'\')+[RrLsCtg])+'\')+[RrLabMCd])+'\')+CONVERT([char](12),[RrFrLn],(0)))+'\')+CONVERT([char](12),[RrToLn],(0)))+'\') persisted
GOGO

/* adding index to RmRt table */
IF not EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[RmRt]') AND name = N'UK_RmRt')
CREATE UNIQUE NONCLUSTERED INDEX [UK_RmRt] ON [dbo].[RmRt] 
(
	[RrAllKeys] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 80) ON [PRIMARY]
GOGO


IF NOT EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('OrdDsg') AND upper(COLUMN_NAME) = upper('OdMasterPrc'))
BEGIN    
	Alter table OrdDsg Add OdMasterPrc float NOT NULL CONSTRAINT DF_OdMasterPrc DEFAULT 0
END

GOGO

IF NOT EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('OrdDsg') AND upper(COLUMN_NAME) = upper('OdTag'))
BEGIN    
	Alter table OrdDsg Add OdTag float NOT NULL CONSTRAINT DF_OdTag DEFAULT 0
END
GOGO

IF NOT EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('OrdDsg') AND upper(COLUMN_NAME) = upper('OdProfit'))
BEGIN    
	Alter table OrdDsg Add OdProfit float NOT NULL CONSTRAINT DF_OdProfit DEFAULT 0
END
GOGO

IF NOT EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
    WHERE  upper(TABLE_NAME) = upper('OrdDsg') AND upper(COLUMN_NAME) = upper('OdProfitPer'))
BEGIN    
	Alter table OrdDsg Add OdProfitPer float NOT NULL CONSTRAINT DF_OdProfitPer DEFAULT 0
END
GOGO

/* Value changed in TcMst table */
Update TcMst set TmInvHdMustYN ='Y' where TmTc='JPR' or TmTc='JCN' 


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
	/******5717 checks missed are taken from 577 and added to 5716 trigger  *****/
	Select @Row=Count('x') From Inserted ins where 
			(Select hHolChkYN From Head where HCd='ZSELF' and HCoCd=BCoCd)='Y' and
			(@Mode='A' or (@Mode='C' and  ins.BOpnDt<>IsNull((Select BOpnDt From Deleted del where 
									ins.BPrtKey=del.BPrtKey and ins.BIdNo=del.BIdNo 
									and ins.BCoCd= del.BCoCd and  ins.BYy= del.BYy 
										and ins.BChr= del.BChr and ins.BNo= del.BNo),ins.BOpnDt))) and  
								(Select top 1 DtHolYN From DtTbl Where DtTag='HOL' and BOpnDt=DtDt and DtCoCd in ('ZZZ',ins.BCoCd) order by DtCoCd asc)='Y'
	If @Row> 0
	begin  
		select @errmsg= 'Cannot Open Bag on Holidays.'  
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
		select @row= count('x') from deleted del where del.BPrtKey=@wCurrPrtn and del.BCls= 'Y' and del.BRepYN ='N' and
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
		end
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


ALTER TRIGGER [dbo].[Fgd_INUPDL] ON [dbo].[Fgd]     
        FOR INSERT,UPDATE,DELETE       
	NOT FOR REPLICATION
AS      
BEGIN      
 DECLARE @ROW INT      
 DECLARE @mode char(1)      
 DECLARE @errmsg varchar(255)      
 Declare @wTcTyp Varchar(3)    
 Declare @wHCd Varchar(8)     
 DECLARE @ReplAgt BIT     
 DECLARE @wCurrPrtn VarChar(1)     
      
 IF @@ROWCOUNT=0   RETURN      
      
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'Fgd' )      
  RETURN      
     
 select @mode=case (select count('x') from inserted) when 0 then 'D'      
       else      
   case (select count('x') from deleted) when 0 then 'A' else 'C' end      
       end       
    
 Select @wHCd= 'ZSELF'    
 Select @wCurrPrtn='C'    
 /*** Storing Tc Type In The Variable @wTcTyp ***/    
 if @mode in ('A','C')    
 Begin    
  select @wTcTyp= max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= FdTc and PSCd=''    
 end    
 if @mode in ('D')    
 Begin    
  select @wTcTyp= max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= FdTc and PSCd=''    
 end    
    
 
 if @mode='A'    
 Begin    
  Select @Row= Count('x') from inserted    
   join Bag on FdBIdNo=BIdNo and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo     
     and FdPrtKey=BPrtKey    
   where FdPrtKey=@wCurrPrtn and FdDt<BOpnDt    
  If @Row > 0    
  begin    
   Select @ErrMsg= 'FG Date Cannot Be Less Than Bag Opening Date'    
   GoTo ErrHndlr    
  end
  --5717
  select @Row= Count('x') from inserted ins where     
   exists(Select 'x' from Fgd Fd1 where     
    Fd1.FdCoCd= ins.FdCoCd and Fd1.FdYy= ins.FdYy and Fd1.FdKey= ins.FdKey     
    and (Fd1.FdCoCd<> ins.FdCoCd or Fd1.FdTc<> ins.FdTc or Fd1.FdYy<> ins.FdYy     
     or Fd1.FdChr<> ins.FdChr or Fd1.FdNo<> ins.FdNo or Fd1.FdSr<> ins.FdSr))     
    If @Row> 0    
  begin    
   Select @ErrMsg= 'Yy-Key Combination Already Exists For Another Voucher Sr.'    
   GoTo ErrHndlr    
  end        
 End     
 --5717
 If @Mode = 'C'    
 BEGIN    
  /******* If Mode = 'C' then Cannot change Key Values,Bag No ********/    
  /* Current Partition */    
  If Update(FdPrtKey) or Update(FdIdNo) or Update(FdCoCd) or Update(FdTc) or Update(Fdyy) or Update(FdChr) or
      Update(FdNo) or Update(FdSr) or Update(FdBIdNo) or Update(FdBYy) or Update(FdBChr) or Update(FdBNo) or
      Update(FdKey) or Update(FdRefYy) or Update(FdRefKey) or Update(FdBYyKey) or Update(FdBm) 
  Begin  
    select @Row=Count('x') from Deleted Del where Del.FdPRtKey=@wCurrPrtn and    
      Not Exists (Select 'x' from Inserted Ins     
      where Ins.FdPrtKey=Del.FdPrtKey and Ins.FdIdNo=Del.FdIdNo and Ins.FdCoCd=Del.FdCoCd     
       and Ins.FdTc=Del.FdTc and Ins.Fdyy=Del.FdYy and Ins.FdChr= Del.FdChr     
       and Ins.FdNo=Del.FdNo and Ins.FdSr=Del.FdSr and Ins.FdBIdNo=Del.FdBIdNo and Ins.FdBYy= Del.FdBYy     
       and Ins.FdBChr= Del.FdBChr and Ins.FdBNo= Del.FdBNo and Ins.FdKey= Del.FdKey     
       and Ins.FdRefYy= Del.FdRefYy and Ins.FdRefKey= Del.FdRefKey     
       and Ins.FdBYyKey= Del.FdBYyKey and Ins.FdBm= Del.FdBm)    
    If @Row > 0    
    Begin        
     Select @ErrMsg= 'Cannot Change Voucher No./Bag No/Key'    
     GoTo ErrHndlr    
    End    
  End   
  --5717
  If Update(FdDespQty) or Update(FdCTBDespQty) or Update(FdDmCd) or Update(FdSfx) or Update(FdDmSz) or Update(FdQty) or
      Update(FdGrWt) or Update(FdExpCmCd) or Update(FdExpOdTc) or Update(FdExpOdYy) or Update(FdExpOdChr) or Update(FdExpOdNo) or
      Update(FdExpOdSr) or Update(FdPlYy) or Update(FdPlChr) or Update(FdPlNo) or Update(FdInTc) or
      Update(FdInYy) or Update(FdInChr) or Update(FdInNo) or Update(FdInExpNo)
  Begin  
    select @Row= Count('x') from Inserted Ins join Deleted Del on Ins.FdIdNo=Del.FdIdNo and Ins.FdCoCd= Del.FdCoCd     
     and Ins.FdTc= Del.FdTc and Ins.Fdyy= Del.FdYy     
     and Ins.FdChr= Del.FdChr and Ins.FdNo= Del.FdNo     
     and Ins.FdSr= Del.FdSr and Ins.FdPRtKey=del.FdPRtKey    
    where Ins.FdPrtKey=@wCurrPrtn and (Del.FdDespQty<> 0 or Del.FdCTBDespQty<> 0)     
     and (Ins.FdDmCd<> Del.FdDmCd or Ins.FdSfx<> Del.FdSfx or Ins.FdDmSz<> Del.FdDmSz or     
      Ins.FdQty<> Del.FdQty or Ins.FdGrWt<> Del.FdGrWt or     
      Ins.FdExpCmCd<> Del.FdExpCmCd or Ins.FdExpOdTc<> Del.FdExpOdTc or Ins.FdExpOdYy<> Del.FdExpOdYy or     
      Ins.FdExpOdChr<> Del.FdExpOdChr or Ins.FdExpOdNo<> Del.FdExpOdNo or Ins.FdExpOdSr<> Del.FdExpOdSr or     
      Ins.FdPlYy<> Del.FdPlYy or  Ins.FdPlChr<> Del.FdPlChr or Ins.FdPlNo<> Del.FdPlNo or Ins.FdInTc<> Del.FdInTc or     
      Ins.FdInYy<> Del.FdInYy or Ins.FdInChr<> Del.FdInChr or Ins.FdInNo<> Del.FdInNo or     
      Ins.FdInExpNo<> Del.FdInExpNo     
         )    
    If @Row > 0    
    Begin        
     Select @ErrMsg= 'Cannot Change Dsg/ Suf/ Sz/ Qty/ Grs Wt/ Exp Det when DespQty<> 0 or CTBDespQty<> 0'    
     GoTo ErrHndlr    
    End    
  EnD
  --5717
  If Update(FdPrdOdTc) or Update(FdPrdOdYy) or Update(FdPrdOdChr) or Update(FdPrdOdNo) or
      Update(FdPrdOdSr) or Update(FdQty) 
  Begin  
    select @Row= Count('x') from Deleted Del where Del.FdPrtKey=@wCurrPrtn and Del.FdBNo<>0 and     
      Not Exists (Select 'x' from Inserted Ins where Ins.FdPrtKey=Del.FdPrtKey and Ins.FdIdNo=Del.FdIdNo     
       and Ins.FdCoCd=Del.FdCoCd     
       and Ins.FdTc=Del.FdTc and Ins.Fdyy=Del.FdYy and Ins.FdChr= Del.FdChr and Ins.FdNo=Del.FdNo     
       and Ins.FdSr=Del.FdSr and Ins.FdPrdOdTc= Del.FdPrdOdTc and Ins.FdPrdOdYy= Del.FdPrdOdYy     
       and Ins.FdPrdOdChr= Del.FdPrdOdChr and Ins.FdPrdOdNo= Del.FdPrdOdNo and Ins.FdPrdOdSr= Del.FdPrdOdSr     
       and Ins.FdQty= Del.FdQty)    
    If @Row > 0    
    Begin        
     Select @ErrMsg= 'Cannot Change Production Fields/Bag Quantity'    
     GoTo ErrHndlr    
    End  
enD	
	--5717
	If Update(FdGrWt) or Update(FdBNo) 
  Begin     
    select @Row= Count('x') from Deleted Del where Del.FdPrtKey=@wCurrPrtn and Del.FdBNo<>0 and     
      Not Exists (Select 'x' from Inserted Ins where Ins.FdPrtKey=Del.FdPrtKey     
       and Ins.FdIdNo=Del.FdIdNo and Ins.FdCoCd= Del.FdCoCd     
       and Ins.FdTc= Del.FdTc and Ins.Fdyy= Del.FdYy     
       and Ins.FdChr= Del.FdChr and Ins.FdNo= Del.FdNo     
       and Ins.FdSr= Del.FdSr and Ins.FdGrWt= Del.FdGrWt)    
    If @Row > 0    
    Begin        
     Select @ErrMsg= 'Cannot Change Gross Weight'    
     GoTo ErrHndlr    
    End    
  End  
    
  

  end
  --5717
  If @mode='A' or (@mode='C' or (Update(FdSr)))
  Begin     
    Select @Row= Count('x') From Inserted Ins Where FdPRtKEy=@wCurrPrtn and FdSr= 0    
    If @Row> 0    
    begin    
       Select @ErrMsg= 'Sr. Cannot be 0'    
       GoTo ErrHndlr    
    end    
  End
  --5717
   If @mode='A' or (@mode='C' or (Update(FdKey)))
  Begin     
    Select @Row= Count('x') From Inserted Ins Where FdPRtKey=@wCurrPrtn and FdKey<= 0    
    If @Row> 0    
    begin    
       Select @ErrMsg= 'Key Cannot be <= 0'    
       GoTo ErrHndlr    
    end    
  End
  --5717  
 if @mode in ('A', 'C') and @wTcTyp= 'FB'    
 Begin   
 
   If @mode='A' or (@mode='C' and (Update(FdDespQty)))
    Begin
      Select @Row= Count('x') From Inserted Ins Where FdPrtKey=@wCurrPrtn and FdDespQty< 0    
      If @Row> 0    
      begin    
         Select @ErrMsg= 'Finished Goods Despatch Qty cannot Be Less Than Zero'    
         GoTo ErrHndlr    
      end    
    End
	--5717
	If @mode='A' or (@mode='C' and (Update(FdDmCd) or Update(FdSfx) or Update(FdDmSz) or Update(FdPrdCmCd)))
    Begin
      select @Row= Count('x') from inserted     
        join Bag on FdPrtKey=BPrtKey and FdBIdNo=BIdNo and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo     
        join OrdMst on OmPrtKey=BPrtKey and OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc=BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo     
       where FdPrtKey=@wCurrPrtn and (FdDmCd<> BOdDmCd or FdSfx<> BOdSfx or FdDmSz<> BOdDmSz or FdPrdCmCd<> OmCmCd)    
        and FdBNo<>0    
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Design/Customer Code Fields Do Not Match With Bag'    
       GoTo ErrHndlr    
      end     
    End
	--5717
	If @mode='A' or (@mode='C' and (Update(FdExpOdTc) or Update(FdExpOdYy) or Update(FdExpOdChr) or Update(FdExpOdNo)
                         or Update(FdExpOdSr) or Update(FdExpCmCd)))
    Begin
      select @Row= Count('x') from inserted where FdPrtKey=@wCurrPrtn and    
        (FdExpOdTc= '' or FdExpOdYy= '' or FdExpOdChr= '' or FdExpOdNo= 0 or FdExpOdSr= 0 or     
        FdExpCmCd= '')    
      If @Row > 0    
      begin     
       Select @ErrMsg= 'Export Order And Export Customer Should Be Entered'    
       GoTo ErrHndlr    
      end    
    End
	--5717
	If @mode='A' or (@mode='C' and (Update(FdExpOdTc) or Update(FdExpOdYy) or Update(FdExpOdChr) or Update(FdExpOdNo)
                         or Update(FdExpOdSr) or Update(FdDmCd)))
    Begin
      Select @row= Count('x') from inserted where FdPrtKey=@wCurrPrtn and    
        (    
        (not exists(select 'x'  from param where PTyp='TC'    
         and PMcd=FdExpOdTc and PScd='' and (PValue='SO' or PValue='OS')))    
        OR    
        (not exists(select 'x' from OrdDsg where FdPrtKey=OdPrtKey and FdCoCd=OdCoCd and     
         FdExpOdTc=OdTc and FdExpOdYy=OdYy and FdExpOdChr=OdChr and     
         FdExpOdNo=OdNo and FdExpOdSr=OdSr and FdDmCd= OdDmCd ))     
        )    
        and (FdExpOdTc<>'' or FdExpOdYy<>'' or FdExpOdChr<>'' or     
         FdExpOdNo<>0 or FdExpOdSr<>0)    
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Exp Ord No does not exist / Dsg Cd MisMatch/ Exp.Tc<> ''SO'' Or ''OS''  '    
       GoTo ErrHndlr    
      end    
    End

	--5717
	If @mode='A' or (@mode='C' and (Update(FdExpOdTc) or Update(FdExpOdYy) or Update(FdExpOdChr) or Update(FdExpOdNo)
                         or Update(FdExpOdSr) or Update(FdExpCmCd)))
    Begin
      select @Row= Count('x') from inserted     
        join OrdMst on OmPrtKey=FdPrtKey and OmCoCd= FdCoCd and OmTc= FdExpOdTc and OmYy= FdExpOdYy and OmChr= FdExpOdChr and OmNo= FdExpOdNo     
       where FdPrtKey=@wCurrPrtn and FdExpCmCd<> OmCmCd     
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Export Customer Code Does Not Match With Exp Ord'    
       GoTo ErrHndlr    
      end    
    End
	--5717
	If @mode='A' or (@mode='C' and (Update(FdExpOdTc) or Update(FdExpOdYy) or Update(FdExpOdChr) or Update(FdExpOdNo)
                         or Update(FdExpOdSr) or Update(FdPrdOdTc) or Update(FdPrdOdYy) or Update(FdPrdOdChr) or Update(FdPrdOdNo)
                         or Update(FdPrdOdSr)))
    Begin
      select @Row= Count('x') from inserted join OrdRm Pr on OrPrtKey=FdPrtKey and OrCoCd= FdCoCd and     
       OrTc= FdExpOdTc and OrYy= FdExpOdYy and OrChr= FdExpOdChr and     
       OrNo= FdExpOdNo and OrSr= FdExpOdSr     
       where FdPrtKey=@wCurrPrtn and not exists(Select 'x' from OrdRm Ex Where OrPrtKey=FdPrtKey and OrCoCd= FdCoCd and     
        OrTc= FdPrdOdTc and OrYy= FdPrdOdYy and OrChr= FdPrdOdChr and     
        OrNo= FdPrdOdNo and OrSr= FdPrdOdSr and Ex.OrRmCd= Pr.OrRmCd)     
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Rm In Export Ord Not Matching With Rm In Prd Ord'    
       GoTo ErrHndlr    
      end  
	  --5717
	  Select @Row= Count('x') From Inserted Where FdPrtKEy=@wCurrPrtn and (FdPrdOdTc <> FdExpOdTc Or     
        FdPrdOdYy <> FdExpOdYy Or FdPrdOdChr <> FdExpOdChr Or     
        FdPrdOdNo <> FdExpOdNo Or FdPrdOdSr <> FdExpOdSr) And     
       Exists    
       (Select 1 From     
        (Select OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd from OrdRm     
        Where OrPrtKey=FdPRtKey and OrCoCd = FdCoCd And OrTc = FdPrdOdTc And OrYy = FdPrdOdYy And     
         OrChr = FdPrdOdChr And OrNo = FdPrdOdNo And OrSr = FdPrdOdSr) PR    
       FULL OUTER JOIN     
        (Select OrCoCd, OrTc, OrYy, OrChr, OrNo, OrSr, OrRmCd from OrdRm     
        Where OrPrtKey=FdPrtKey and OrCoCd = FdCoCd And OrTc = FdExpOdTc And OrYy = FdExpOdYy And     
         OrChr = FdExpOdChr And OrNo = FdExpOdNo And OrSr = FdExpOdSr) ER    
       ON PR.OrRmCd = ER.OrRmCd    
       WHERE ((PR.OrRmCd is Null) Or (ER.OrRmCd is Null))    
       )    
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Rm In Export Order Not Matching With Rm In Production Order'    
       GoTo ErrHndlr    
      end    
    End    
	--5717
	If @mode='A' or (@mode='C' and (Update(FdDespQty) or Update(FdQty)))
    Begin
      Select @Row=Count('x') From Inserted Where FdPrtKey=@wCurrPrtn and FdDespQty > FdQty    
        If @Row>0    
      begin    
       Select @ErrMsg= 'Desp Qty should be<= Quantity'    
         GoTo ErrHndlr    
      end    
    End
	--5717
	If @mode='A' or (@mode='C' and (Update(FdRefYy) or Update(FdYy) or Update(FdRefKey) or Update(FdKey)))
    Begin
      Select @Row=Count('x') From Inserted Where FdPrtKEy=@wCurrPrtn and (FdRefYy<> FdYy  or FdRefKey<> FdKey)    
        If @Row>0    
      begin    
       Select @ErrMsg= 'Reference Yy/Key Should Be Equal To Yy/Key'    
         GoTo ErrHndlr    
      end    
    End
END

 --5717
 If @wTcTyp in ('FWB', 'FWR')
  Begin
    If Update(FdDmCd) or Update(FdSfx) or Update(FdDmSz) or Update(FdQty) or
        Update(FdGrWt) or Update(FdExpCmCd) or Update(FdExpOdTc) or Update(FdExpOdYy) or Update(FdExpOdChr) or 
        Update(FdExpOdNo) or Update(FdExpOdSr) or Update(FdPlYy) or Update(FdPlChr) or Update(FdPlNo) or 
        Update(FdInTc) or Update(FdInYy) or Update(FdInChr) or Update(FdInNo) or Update(FdInExpNo) or
        Update(FdRefYy) or Update(FdRefKey) 
    Begin  
      select @Row= Count('x') from Inserted Ins join Deleted Del on Ins.FdPrtKey=Del.FdPrtKey     
       and Ins.FdIdNo=Del.FdIdNo and Ins.FdCoCd= Del.FdCoCd     
       and Ins.FdTc= Del.FdTc and Ins.Fdyy= Del.FdYy     
       and Ins.FdChr= Del.FdChr and Ins.FdNo= Del.FdNo     
       and Ins.FdSr= Del.FdSr     
      where @wTcTyp in ('FWB', 'FWR') and Ins.FdPrtKey=@wCurrPrtn and (Ins.FdDmCd<> Del.FdDmCd or Ins.FdSfx<> Del.FdSfx or Ins.FdDmSz<> Del.FdDmSz or     
         Ins.FdQty<> Del.FdQty or Ins.FdGrWt<> Del.FdGrWt or     
         Ins.FdExpCmCd<> Del.FdExpCmCd or Ins.FdExpOdTc<> Del.FdExpOdTc or Ins.FdExpOdYy<> Del.FdExpOdYy or     
         Ins.FdExpOdChr<> Del.FdExpOdChr or Ins.FdExpOdNo<> Del.FdExpOdNo or Ins.FdExpOdSr<> Del.FdExpOdSr or     
         Ins.FdPlYy<> Del.FdPlYy or  Ins.FdPlChr<> Del.FdPlChr or Ins.FdPlNo<> Del.FdPlNo or Ins.FdInTc<> Del.FdInTc or     
         Ins.FdInYy<> Del.FdInYy or Ins.FdInChr<> Del.FdInChr or Ins.FdInNo<> Del.FdInNo or     
         Ins.FdInExpNo<> Del.FdInExpNo or Ins.FdRefYy<> Del.FdRefYy or Ins.FdRefKey<> Del.FdRefKey)     
      If @Row > 0    
      Begin        
       Select @ErrMsg= 'Cannot Change Dsg/ Suf/ Sz/ Qty/ Grs Wt/ Exp Det/ Reference Yy, Key'    
       GoTo ErrHndlr    
      End    
    End
  End

     
 If @mode= 'A' And @wTcTyp= 'FB'    
 Begin    
  Select @Row= Count('x') from inserted    
   Where FdPrtKey=@wCurrPrtn     
    and not exists(select 'x' from Bag where FdPrtKey=BPrtKey and FdBIdNo=BIdNo     
     and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo and BCls= 'N'     
     and BPtQty= 0 and BRjQty= 0 and BRjPtQty= 0 and BLoc<>'PFG')    
  If @Row > 0    
  Begin     
   Select @ErrMsg= 'Invalid Bag No/ Bag Closed/ Bag Already In Fg/ Bag Rej Or Part Qty <> 0'    
   GoTo ErrHndlr    
  End    
 End      
    
 If @mode= 'C'    
  begin    
  select @Row= Count('x') from Deleted del where FdPrtKey=@wCurrPrtn and     
       Exists(Select 'x' from Fgd Fgd1 where Fgd1.FdPrtKey=del.FdPrtKey and Fgd1.FdBIdNo=del.FdBIdNo     
    and Fgd1.FdCoCd= del.FdCoCd and Fgd1.FdBYy= del.FdBYy and Fgd1.FdBChr= del.FdBChr     
    and Fgd1.FdBNo= del.FdBNo and     
       ((Fgd1.FdYy > del.FdYy)     
   or     
   (Fgd1.FdYy = del.FdYy and Fgd1.FdKey> del.FdKey)    
              ))      
  If @Row> 0    
  Begin        
   Select @ErrMsg= 'Cannot Change As Not The Last Transaction Of The Bag'    
   GoTo ErrHndlr    
  End    
 End     
    
 if @mode in ('A', 'C') and @wTcTyp= 'FWB'    
 Begin    
  if @mode='A'    
  Begin    
   select @Row= Count('x') from inserted    
    where FdPrtKey=@wCurrPrtn and not exists(select 'x' from Bag where FdPrtKey=BPrtKey and FdBIdNo=BIdNo     
     and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo     
     and BCls= 'N' and BLoc= 'PFG')    
   If @Row > 0    
   begin     
    Select @ErrMsg= 'Invalid Bag No/ Bag Closed/ Bag Not In Fg'    
    GoTo ErrHndlr    
   end    
  End    
 End     
 If (@mode='D')    
 Begin    
	declare @wBMaxTxnDt date
	Select @wBMaxTxnDt = 
					    case when del.FdTc='FB' then 
							  isnull((select max(FdDt) from FgD where fgd.fdbyykey=del.fdbyykey),
						isnull((Select max(ForDt.TdDt) from Txnd ForDt where ForDt.TdBIdNo=Bag.BIdNo
															and ForDt.TdYyKey=Del.FdByyKey 
															),BOpnDt))
					   when del.fdtc='FWB' then 
					   (select max(ForPrevDt.FdDt) from FgD 
									join Fgd ForPrevDt on 
									ForPrevDt.FdYy=Del.FdRefYy and ForPrevDt.fdkey=del.FdRefKey)
					   else BMaxTxnDt end
  from Bag
  join Deleted del on
  Bag.BIdNo=del.FdBIdNo
 End      
    
 if @mode in ('D') and @wTcTyp in ('FB', 'FR','FWB')    
 Begin    
  Delete FgRm From deleted,FgRm where FdPrtKey=FrPrtKey and FdIdNo=FrFdIdNo and FdCoCd = FrCoCd and     
   FdTc = FrTc and FdYy = FrYy and FdChr = FrChr and FdNo = FrNo and FdSr= FrSr     

  if @wTcTyp='FWB'
	  Update Bag Set
	  BMaxTxnDt=@wBMaxTxnDt,
	  BLoc=   Isnull(F.FgToBLoc, BOpnLoc),     
	  BRecvDt= Isnull(F.FgDt, BOpnDt),
	  BPrdStg= '',     
	  BFgSubLoc=  ''
	from Deleted D
	join Bag on D.FdBIdNo = Bag.BIdNo 
	left outer join Fgd ForMatchingFG on ForMatchingFG.FdYy=D.FdRefYY and ForMatchingFG.fdkey=D.FdRefKey
	left outer join Fg F on ForMatchingFg.FdFgIdNo =F.FgIdNo
  else
  Update Bag Set
  BMaxTxnDt=@wBMaxTxnDt,
  BLoc=   Isnull(F.FgToBLoc, IsNull(TdToBLoc, BOpnLoc)),     
  BRecvDt= Isnull(F.FgDt, Isnull(TdDt, BOpnDt)),
  BPrdStg= '',     
  BFgSubLoc=  ''     
  From Deleted D 
   Left Outer Join Fgd Fd On Fd.FdBIdNo=D.FdBIdNo and ((Fd.FdBYyKey= D.FdBYyKey)/* or 
     Not Exists(Select 'x' from Txnd Td Where Td.TdSrNo= 0 and td.TdBIdNo=Fd.FdBIdNo and td.TdCoCd= Fd.FdCoCd 
      and td.TdBYy= Fd.FdBYy and td.TdBChr= Fd.FdBChr and td.TdBNo= Fd.FdBNo and td.TdPrtKey=Fd.FdPrtKey and
     TdYyKey>Fd.FdBYyKey and TdToBLoc<>'')*/)
     and Fd.FdYy+ Str(Fd.FdKey)=       
     (Select max(Fd1.FdYy+ Str(Fd1.FdKey))     
     From Fgd Fd1 where Fd1.FdTc= 'FWB'       
      and Fd1.FdBIdNo=Fd.FdBIdNo) 
   Left Outer join Fg F on F.FgIdNo=Fd.FdFgIdNo 
   Left Outer Join Txnd TdDt on TdDt.TdBIdNo=D.FdBIdNo 
     and TdDt.TdSrNo= 0 and TdDt.TdYyKey=     
     (
	 Select max(td.TDYyKey) from Txnd td where     
     Td.TdSrNo= 0 and Td.TdToBLoc<> '' and td.TdBIdNo=D.FdBIdNo)   
  Where BIdNo= D.FdBIdNo

  --Update Bag Set
  --BLoc=   Isnull(F.FgToBLoc, IsNull(TdToBLoc, BOpnLoc)),     
  --BRecvDt= Isnull(F.FgDt, Isnull(TdDt, BOpnDt)),
  --BPrdStg= IsNull((Select PMCd from Param P1 where PTyp= 'PRDSTG'     
  --           and PNum= IsNull((Select max(P2.PNum) from Txnd TdLoc     
  --              join Head on HCoCd= TdCoCd and HCd= 'ZSELF'     
  --              join Loc on LocCoCd= TdLoc.TdCoCd and LocTyp= 'P' and LocCd= TdLoc.TdByLoc     
  --              join Param P2 on P2.PTyp= 'PRDSTG' and P2.PMCd= LocPrdStg     
  --              where D.FdBIdNo=TdLoc.TdBIdNo and D.FdCoCd= TdLoc.TdCoCd and D.FdBYy= TdLoc.TdBYy     
  --              and D.FdBChr= TdLoc.TdBChr and D.FdBNo= TdLoc.TdBNo and D.FdPrtKey=TdLoc.TdPrtKey    
  --              and TdLoc.TdSrNo= 0 and TdLoc.TdByLoc<> '' and TdLoc.TdPrdYN= 'Y' and hPpcYN= 'Y'     
  --             ), 0)     
  --           ), ''),     
  --BFgSubLoc=  ''     
  --From Deleted D 
  -- Left Outer Join Fgd Fd On Fd.FdBIdNo=D.FdBIdNo and Fd.FdCoCd= D.FdCoCd and Fd.FdBYy= D.FdBYy     
  --   and Fd.FdBChr= D.FdBChr and Fd.FdBNo= D.FdBNo and Fd.FdPrtKey=D.FdPrtKey    
  --   and ((Fd.FdBYyKey= D.FdBYyKey) or 
  --   Not Exists(Select 'x' from Txnd Td Where Td.TdSrNo= 0 and td.TdBIdNo=Fd.FdBIdNo and td.TdCoCd= Fd.FdCoCd 
  --    and td.TdBYy= Fd.FdBYy and td.TdBChr= Fd.FdBChr and td.TdBNo= Fd.FdBNo and td.TdPrtKey=Fd.FdPrtKey and
  --   TdYyKey>Fd.FdBYyKey and TdToBLoc<>''))
  --   and Fd.FdYy+ Str(Fd.FdKey)=       
  --   (Select max(Fd1.FdYy+ Str(Fd1.FdKey))     
  --   From Fgd Fd1 where Fd1.FdTc in (Select PMCd from Param where PTyp= 'TC' and PValue=(Case when D.FdTc In ('FB','FR') then 'FWB' Else 'FB' End))      
  --    and Fd1.FdBIdNo=Fd.FdBIdNo and Fd1.FdCoCd= Fd.FdCoCd and Fd1.FdBYy= Fd.FdBYy     
  --    and Fd1.FdBChr= Fd.FdBChr and Fd1.FdBNo= Fd.FdBNo and Fd1.FdPrtKey=Fd.FdPrtKey) 
  -- Left Outer join Fg F on F.FgIdNo=Fd.FdFgIdNo and F.FgCoCd= Fd.FdCoCd     
  --   and F.FgTc= Fd.FdTc and F.FgYy= Fd.FdYy     
  --   and F.FgChr= Fd.FdChr and F.FgNo= Fd.FdNo and F.FgPrtKey=Fd.FdPrtKey          
  -- Left Outer Join Txnd TdDt on TdDt.TdBIdNo=D.FdBIdNo and TdDt.TdCoCd= D.FdCoCd and TdDt.TdBYy= D.FdBYy     
  --   and TdDt.TdBChr= D.FdBChr and TdDt.TdBNo= D.FdBNo and TdDt.TdPrtKey=D.FdPrtKey    
  --   and TdDt.TdSrNo= 0 and TdDt.TdYyKey=     
  --   (select max(td.TDYyKey) from Txnd td where     
  --   Td.TdSrNo= 0 and Td.TdToBLoc<> '' and td.TdBIdNo=D.FdBIdNo    
  --   and td.TdCoCd= D.FdCoCd and td.TdBYy= D.FdBYy     
  --   and td.TdBChr= D.FdBChr and td.TdBNo= D.FdBNo and td.TdPrtKey=D.FdPrtKey)   
  --Where BIdNo= D.FdBIdNo and BCoCd= D.FdCoCd and BYy= D.FdBYy and BChr= D.FdBChr and BNo= D.FdBNo 
  --      and BPrtKey=D.FdPrtKey and D.FdPrtKey=@wCurrPrtn
    
  update OrdDsg set     
  ModUsr='MIC', OdFgQty= OdFgQty- FdQty     
  from deleted where FdPrtKey=@wCurrPrtn and OdCoCd= FdCoCd and OdTc= FdPrdOdTc and OdYy= FdPrdOdYy and     
   OdChr= FdPrdOdChr and OdNo=FdPrdOdNo and OdSr= FdPrdOdSr and OdPrtKey=FdPrtKey and @wTcTyp<>'FWB'   
 End     
    
 if @mode in ('A') and @wTcTyp in ('FB', 'FR')    
 Begin    
  update Bag set     
  BLoc= (select FgToBLoc from Fg join inserted on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd     
   and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo and FgPrtKey=FdPrtKey),     
  BRecvDt= (select FgDt from Fg join inserted on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd     
   and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo and FgPRtKey=FdPrtKey),     
  BPrdStg= '',     
  BFgSubLoc= FdSubLoc     
  from inserted where FdPrtKey=@wCurrPrtn and BIdNo=FdBIdNo and BCoCd= FdCoCd and BYy= FdBYy and BChr=FdBChr and BNo=FdBNo and BPrtKey=FdPrtKey    
    
  update OrdDsg set     
  ModUsr='MIC',OdFgQty= OdFgQty+ FdQty     
  from inserted where FdPRtkey=@wCurrPrtn and OdCoCd= FdCoCd and OdTc= FdPrdOdTc and OdYy= FdPrdOdYy and     
   OdChr= FdPrdOdChr and OdNo=FdPrdOdNo and OdSr= FdPrdOdSr and OdPrtKey=FdPrtKey    
 End    
    
 /*********** urmi 2.04 *****************/    
 if @mode = 'C' and @wTcTyp = 'FB'    
 begin    
  If Update(FdSubLoc)
  Begin  
    update Bag set     
     BFgSubLoc= IsNull(ins.FdSubLoc, '')     
     From inserted ins     
     where ins.FdPrtKey=@wCurrPrtn and BIdNo=ins.FdBIdNo and BCoCd= ins.FdCoCd and BYy= ins.FdBYy and BChr= ins.FdBChr     
      and BNo= ins.FdBNo and BPrtKey=ins.FdPrtKey and    
      exists(Select 'x' from deleted del where del.FdIdNo=ins.FdIdNo and    
       del.FdCoCd= ins.FdCoCd and del.FdTc= ins.FdTc     
       and del.FdYy= ins.FdYy and del.FdChr= ins.FdChr     
       and del.FdNo= ins.FdNo and del.FdSr= ins.FdSr and del.FdPrtKey=ins.FdPrtKey    
       and ins.FdSubLoc<> del.FdSubLoc)     
  End       
 end
 
 if @mode = 'A'     
 Begin    
   Update Bag     
   Set BMaxTxnDt = FdDt    
   From Inserted        
    Join Bag On BPrtKey=FdPrtKey and BIdNo=FdBIdNo and BCoCd=FdCoCd and BYy=FdBYy and BChr=FdBChr and BNo=FdBNo    
   Where FdPrtKey=@wCurrPrtn and Fddt > BMaxTxnDt       
 End    
    
 if @mode in ('D') and @wTcTyp in ('FWB', 'FWR')    
 Begin    
  select @Row= Count('x') from Deleted del where @wTcTyp= 'FWB' and del.FdPrtKey=@wCurrPrtn and del.FdBNo<> 0 and     
   exists(Select 'x' from Fgd Fg where Fg.FdCoCd= del.FdCoCd and Fg.FdBIdNo=del.FdBIdNo and     
    Fg.FdBYy= del.FdBYy and Fg.FdBChr= del.FdBChr and Fg.FdBNo= del.FdBNo and Fg.FdPrtKey=del.FdPrtKey and     
    (
	(Fg.FdYy> del.FdYy)     
     or     
     (Fg.FdYy= del.FdYy and Fg.FdKey> del.FdKey)     
    )    
   )
   
  If @Row> 0    
  begin        
   Select @ErrMsg= 'Cannot Delete As Not The Last Fg Transaction Of The Bag'    
   GoTo ErrHndlr    
  end    
    
  select @Row= Count('x') from Deleted join Bag on BIdNo=FdBIdNo and BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr     
    and BNo= FdBNo and BPrtKey=FdPrtKey    
   where @wTcTyp= 'FWB' and FdPrtKey=@wCurrPrtn and FdBYyKey<> BLstYyKey     
  If @Row > 0    
  begin        
   Select @ErrMsg= 'Cannot Delete As Some Transaction Of The Bag Has Taken Place After Backward Movement'    
   GoTo ErrHndlr    
  end    
    
  Delete FgRm From deleted,FgRm where FdPrtKey=@wCurrPrtn and FdIdNo=FrFdIdNo and FdCoCd = FrCoCd and FdTc = FrTc     
   and FdYy = FrYy and FdChr = FrChr and FdNo = FrNo and     FdSr= FrSr and FdPrtKey=FrPrtKey    
    
  /******* Update BLoc = 'PFG' ******/     
  update Bag set    
  BLoc= 'PFG',     
  BPrdStg= (case when IsNull((Select max(hPpcYN) from Head where HCoCd= FdCoCd and HCd= @wHCd), '') = 'Y' then (Select PMCd from Param where PTyp= 'PRDSTG' and PNum= 5) else '' end)     
  from deleted where BPrtKey=deleted.FdPrtKey and BIdNo=deleted.FdBIdNo and BCoCd= deleted.FdCoCd and BYy= deleted.FdBYy     
   and BChr= deleted.FdBChr and BNo= deleted.FdBNo     
    
  if @mode = 'D' and @wTcTyp = 'FWB'    
  begin    
   update Bag set BFgSubLoc= IsNull(    
    (Select FmdToFgSubLoc from Fmd Fmd1 where     
         Fmd1.FmdCoCd= del.FdCoCd and Fmd1.FmdBYy= del.FdBYy     
     and Fmd1.FmdBChr= del.FdBChr and Fmd1.FmdBNo= del.FdBNo and Fmd1.FmdPrtKey=del.FdPrtKey    
     and Fmd1.FmdYy= (Select max(Fmd2.FmdYy) from Fmd Fmd2     
       where Fmd2.FmdPrtKey=fmd1.FmdPrtKey and Fmd2.FmdCoCd= fmd1.FmdCoCd     
       and Fmd2.FmdBYy= Fmd1.FmdBYy     
       and Fmd2.FmdBChr= Fmd1.FmdBChr     
       and Fmd2.FmdBNo= Fmd1.FmdBNo)     
     and Fmd1.FmdKey= (Select max(Fmd2.FmdKey) from Fmd Fmd2     
       where Fmd2.FmdPrtKey=Fmd1.FmdPrtKey and Fmd2.FmdCoCd= Fmd1.FmdCoCd     
       and Fmd2.FmdBYy= Fmd1.FmdBYy     
       and Fmd2.FmdBChr= Fmd1.FmdBChr     
       and Fmd2.FmdBNo= Fmd1.FmdBNo    
       and Fmd2.FmdYy= Fmd1.FmdYy)    
    ),     
    (Select FdSubLoc from Fgd Fgd1 where Fgd1.FdBIdNo=del.FdBIdNo and    
         Fgd1.FdCoCd= del.FdCoCd and Fgd1.FdBYy= del.FdBYy and    
      Fgd1.FdBChr= del.FdBChr and Fgd1.FdBNo= del.FdBNo and Fgd1.FdPrtKey=del.FdPrtKey    
     and Fgd1.FdYy= (Select max(Fgd2.FdYy) from Fgd Fgd2     
       where Fgd2.FdPrtKey=Fgd1.FdPrtKey and Fgd2.FdBIdNo=Fgd1.FdBIdNo and Fgd2.FdCoCd= Fgd1.FdCoCd     
       and Fgd2.FdBYy= Fgd1.FdBYy     
       and Fgd2.FdBChr= Fgd1.FdBChr     
       and Fgd2.FdBNo= Fgd1.FdBNo)     
     and Fgd1.FdKey= (Select max(Fgd2.FdKey) from Fgd Fgd2     
       where Fgd2.FdPrtKey=fgd1.FdPrtKey and Fgd2.FdBIdNo=fgd1.FdBIdNo and Fgd2.FdCoCd= fgd1.FdCoCd     
       and Fgd2.FdBYy= Fgd1.FdBYy     
       and Fgd2.FdBChr= Fgd1.FdBChr     
       and Fgd2.FdBNo= Fgd1.FdBNo    
       and Fgd2.FdYy= Fgd1.FdYy)    
    ))      
   From deleted del where BPrtKey=del.FdPrtKey and BIdNo=del.FdBIdNo and BCoCd= del.FdCoCd and    
    BYy= del.FdBYy and BChr= del.FdBChr and BNo= del.FdBNo    
  end    
  /*********** urmi 2.04 *****************/     
    
  /******* update FdDespQty of Fgd Base Record ******/     
  /* Current Partition */    
  update Fgd set     
  Fgd.FdDespQty= Fgd.FdDespQty- IsNull((Select sum(del1.FdQty) from deleted del1 where     
   del1.FdPrtKey=@wCurrPrtn and del1.FdPrtKey=del.FdPrtKey and del1.FdIdNo=del.FdIdNo and del1.FdCoCd= del.FdCoCd and del1.FdTc= del.FdTc and del1.FdYy= del.FdYy     
   and del1.FdChr= del.FdChr and del1.FdNo= del.FdNo and del1.FdSr= del.FdSr ), 0)     
  from deleted del where Fgd.FdPRtKey=del.FdPrtKey and Fgd.FdCoCd= del.FdCoCd and Fgd.FdYy= del.FdRefYy and Fgd.FdKey= del.FdRefKey and Fgd.FdBM= 'B'     
    
  /******* update OdFgQty of OrdDsg Record corresponding to the PrdOrder ******/     
  /* Current Partition */    
  update OrdDsg set     
  ModUsr='MIC',
  OdFgQty= OdFgQty+ FdQty     
  from deleted where OdPrtKey=FdPrtKey and OdCoCd= FdCoCd and OdTc= FdPrdOdTc and OdYy= FdPrdOdYy and     
   OdChr= FdPrdOdChr and OdNo=FdPrdOdNo and OdSr= FdPrdOdSr     
 End     
    
 if @mode in ('A') and @wTcTyp in ('FWB', 'FWR')    
 Begin    
  update Bag set     
  BLoc= (select FgToBLoc from Fg join inserted on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd     
         and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo and FgPrtKey=FdPrtKey),     
 /***** urmi 2.06 BRecvDt *******/    
  BRecvDt= (select FgDt from Fg join inserted on FgIdNo=FdFgIdNo and FgCoCd= FdCoCd     
            and FgTc= FdTc and FgYy= FdYy and FgChr= FdChr and FgNo= FdNo and FgPrtKey=FdPrtKey),     
 /*******************************/    
  BPrdStg= IsNull((Select PMCd from Param P1 where PTyp= 'PRDSTG'     
   and PNum= IsNull((Select max(P2.PNum) from Txnd TdLoc     
      join Head on HCoCd= TdCoCd and HCd= @wHCd     
      join Loc on LocCoCd= TdLoc.TdCoCd and LocTyp= 'P' and LocCd= TdLoc.TdByLoc     
      join Param P2 on P2.PTyp= 'PRDSTG' and P2.PMCd= LocPrdStg     
      where FdBIdNo=TdLoc.TdBIdNo and FdCoCd= TdLoc.TdCoCd and FdBYy= TdLoc.TdBYy     
      and FdBChr= TdLoc.TdBChr and FdBNo= TdLoc.TdBNo and FdPrtKey=TdLoc.TdPRtKey    
      and TdLoc.TdSrNo= 0 and TdLoc.TdByLoc<> '' and TdLoc.TdPrdYN= 'Y' and hPpcYN= 'Y'     
     ), 0)     
   ), ''),     
  BFgSubLoc= ''     
  from inserted where FdPrtKey=@wCurrPrtn and BPrtKey=FdPrtKey and BIdNo=FdBIdNo and BCoCd= FdCoCd and BYy= FdBYy and BChr=FdBChr     
   and BNo=FdBNo     
    
  /******* update FdDespQty of Fgd ******/    
  /* Current Partition */    
  update Fgd set     
  Fgd.FdDespQty= Fgd.FdDespQty+ IsNull((Select sum(ins1.FdQty) from inserted ins1 where     
   ins1.FdPrtKey=ins.FdPrtKey and ins1.FdIdNo=ins.FdIdNo and ins1.FdCoCd= ins.FdCoCd and ins1.FdTc= ins.FdTc and ins1.FdYy= ins.FdYy     
   and ins1.FdChr= ins.FdChr and ins1.FdNo= ins.FdNo and ins1.FdSr= ins.FdSr), 0)     
  from inserted ins where ins.FdPrtKey=@wCurrPrtn and Fgd.FdPrtKey=ins.FdPrtKey and Fgd.FdCoCd= ins.FdCoCd and Fgd.FdYy= ins.FdRefYy and Fgd.FdKey= ins.FdRefKey and Fgd.FdBM= 'B'     
    
  update OrdDsg set     
  ModUsr='MIC',
  OdFgQty= OdFgQty- FdQty     
  from inserted where FdPrtKey=@wCurrPrtn and OdPrtKey=FdPrtKey and OdCoCd= FdCoCd and OdTc= FdPrdOdTc and OdYy= FdPrdOdYy and OdChr= FdPrdOdChr and OdNo=FdPrdOdNo and     
   OdSr= FdPrdOdSr     
 End     
    
--End    
    
-- Normal Trigger Begins here    
    
 /******** Only One TcTyp records can be Inserted at a time.********/    
 /* Partition key not required */    
 select @Row= Count(distinct(PValue)) from inserted Join param On    
  PTyp= 'TC' and PMCd= FdTc and PScd= ''    
 If @Row > 1    
 BEGIN    
  Select @ErrMsg= 'Transactions Of Only One TcType Can Be Acted On'    
    GoTo ErrHndlr    
 END    
    
 /******** Only One TcTyp records can be Deleted at a time. *********/    
 /* Partition key not required */     
 select @Row= Count(distinct(PValue)) from deleted Join param on    
  PTyp= 'TC' and PMcd= FdTc and PSCd=''    
 If @Row > 1    
 BEGIN    
  Select @ErrMsg= 'Transactions Of Only One TcType Can Be Acted On'    
  GoTo ErrHndlr    
 END    
    
 /*** Storing Tc Type In The Variable @wTcTyp ***/    
 if @mode in ('A','C')    
 /* Partition Key not required */    
 Begin    
  select @wTcTyp= max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= FdTc and PSCd=''    
 end    
 if @mode in ('D')    
 /* Partition Key not required */    
 Begin    
  select @wTcTyp= max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= FdTc and PSCd=''    
 end    
    
 If @Mode= 'A'    
 Begin    
 -- check how the checks are 
  Select @Row= Count('x') from inserted ins where      
       exists(Select 'x' from Fgd Fgd1 where Fgd1.FdBIdNo=ins.FdBIdNo and Fgd1.FdCoCd= ins.FdCoCd and     
       Fgd1.FdBYy= ins.FdBYy and Fgd1.FdBChr= ins.FdBChr and     
       Fgd1.FdBNo= ins.FdBNo and Fgd1.FdPrtKey=ins.FdPrtKey and     
       ((Fgd1.FdYy > ins.FdYy)     
   or     
   (Fgd1.FdYy = ins.FdYy and Fgd1.FdKey> ins.FdKey)     
              ))      
  If @Row> 0    
  begin        
   Select @ErrMsg= 'Cannot Insert in Fgd As Voucher Year Less Than The Last Transaction Of The Bag'    
   GoTo ErrHndlr    
  end    
 End     
    
    
 if @mode in ('A', 'C')    
 Begin    
  /******* Valid FdTc,FdYy,FdChr, FdNo combination ********/    
  /* Current Partition */    
  If @mode='A' or (@mode='C' or (Update(FdPrtKey) or Update(FdFgIdNo) or Update(FdCoCd) or Update(FdTc)
                       or Update(FdYy) or Update(FdChr)  or Update(FdNo)))
  Begin     
    Select @Row= Count('x') From Inserted Ins Where Ins.FdPrtKEy=@wCurrPrtn and    
       Not Exists (Select 'x' From Fg Where FgPrtKey=Ins.FdPrtKey and FgIdNo=Ins.FdFgIdNo and FgCoCd=Ins.FdCoCd And     
          FgTc=Ins.FdTc And FgYy=Ins.FdYy And  FgChr=Ins.FdChr And    
              FgNo=Ins.FdNo)    
      If @Row> 0    
     begin    
     Select @ErrMsg= 'Voucher No. Does Not Exist'    
     GoTo ErrHndlr    
      end    
  End      

  If @mode='A' 
  Begin
    Select @Row=Count('x') from Inserted     
       Join Bag On FdPrtKey=BPrtKey and FdBIdNo=BIdNo and FdCoCd= BCoCd and    
         FdByy= BYy and FdBChr= BChr and FdBNo= BNo    
      Where @mode='A' and Fddt<BMaxTxnDt    
    If @Row > 0    
    Begin    
     Select @ErrMsg= 'A Transaction of Later Date Exists For this Bag'    
     GoTo ErrHndlr    
    end
  End     
 End    
    
 if @mode in ('A', 'C') and @wTcTyp= 'FB'    
 Begin    
  If @mode='A' or (@mode='C' or (Update(FdBChr)))
  Begin     
     select @Row= count('x') from Inserted where FdPRtKey=@wCurrPrtn and    
      Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= FdBChr and PSCd= '' and PValue<> 'P')     
     If @Row> 0    
        begin     
      Select @ErrMsg= 'Only Primary Bags Can Be Taken To FG'    
      GoTo ErrHndlr    
     end     
  End    

 -- take it to front end 
 If (@Mode='C' and Not Exists (Select * From Inserted Ins     
   Join Deleted Del On Del.FdPrtKey=ins.FdPrtKey and Del.FdIdNo=ins.FdIdNo and Del.FdCoCd=ins.FdCocd     
   and Del.FdTc=ins.FdTc and Del.FdYy=ins.FdYy and del.FdChr=ins.FdChr and     
   del.FdNo=ins.FdNo and Del.FdSr=ins.FdSr     
   Where del.FdDespQty<>0 or ins.FdDespQty<>0 Or    
    del.FdCTBDespQty<>0 Or ins.FdCTBDespQty<>0 )) Or @Mode='A'    
 Begin      
 --5717
  select @Row=Count('X')  from Inserted Join     
    OrdDsg on OdPrtKey=FdPrtKey and OdCocd= FdCoCd and OdTc= FdPrdOdTc and OdYy= FdPrdOdYy and OdChr=FdPrdOdChr     
      and OdNo= FdPrdOdNo and OdSr= FdPrdOdSr     
  where  ((round(OdMinWt, 3)<> 0 and     
  round((Select sum(TdRmWt* (case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1     
      when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1     
      when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1     
      when  TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end)/ (case when RmCtg in ('C', 'D') then 5 else 1 end) )      
      From Txnd join RmMst on TdRmCd= RmCd and TdPrtKey=RmPrtKey     
      Where /* RmCtg in ('G','P','X')  */    
      ','+ (Select PDesc225 From Param where PTyp= 'CPYRT' and PmCd= OdGmChk)+ ',' like '%,'+ RmCtg+ ',%'     
   and TdBIdNo=FdBIdNo and TdCoCd= FdCoCd and TdBYy= FdBYy and TdBChr= FdBChr and TdBNo= FdBNo and TdPrtKey=FdPrtKey    
   Group By TdCoCd, TdBYy, TdBChr, TdBNo ), 3)< round((OdMinWt* FdQty), 3)) Or    
  (round(OdMaxWt, 3)<> 0 and round((Select sum(TdRmWt*(case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1     
      when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1     
      when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1     
      when  TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end)/ (case when RmCtg in ('C', 'D') then 5 else 1 end)  )      
      From Txnd join RmMst on TdRmCd= RmCd and TdPrtKey=RmPrtKey    
      Where /* RmCtg in ('G','P','X')  */    
      ','+ (Select PDesc225 From Param where PTyp= 'CPYRT' and PmCd= OdGmChk)+ ',' like '%,'+ RmCtg+ ',%'     
   and TdBIdNo=FdBIdNo and TdCoCd= FdCoCd and TdBYy= FdBYy and TdBChr= FdBChr and TdBNo= FdBNo and TdPrtKey=FdPrtKey    
   Group By TdCoCd, TdBYy, TdBChr, TdBNo), 3)> round((OdMaxWt* FdQty), 3)))    
            If @Row > 0    
     begin     
   Select @ErrMsg= 'Gross Gram Wt Is not within the Specified Min & Max Order Gross Wt'    
   /* Gross Wt of Gold,Platinum,Accessories Is not within the Specified Min & Max Order Gross Wt */    
   GoTo ErrHndlr    
  end     

  Select  @Row= Count('x')  from Inserted  where FdPrtKey=@wCurrPrtn and     
            exists ( Select OrRmCd from OrdRm     
        join Bag on BPrtKey=OrPrtKey and BOdIdNo=OrOdIdNo and BCoCd=OrCoCd and BOdTc=OrTc and BOdYy=OrYy    
          and BOdChr=OrChr and BOdNo=OrNo and BOdSr=OrSr    
      where BOdTc='SO' and (select hFgRmEqOrdRm from Head where HCoCd=FdCoCd     
      and HCd=@wHCd)='Y' and    
     (select vPValue3 from vParam where vPTyp='CHR'    
      and vPCoCd=OrCoCd and vPMCd=OrTc and vPSCd=OrChr)='N'  and    
    BPrtKey=FdPrtKey and BIdNo=FdBIdNo and BCoCd= FdCocd and BYy= FdBYy and BChr= FdBChr     
    and BNo= FdBNo and convert(Decimal(10,3),OrPrdWt)>0.001     
    and OrRmCd not in( Select distinct TdRmCd From Txnd      
                   Where TdRmCd <> '' and TdBIdNo=FdBIdNo and TdCoCd= FdCoCd and TdBYy= FdByy    
           and TdBChr= FdBChr and TdBNo= FdBNo and TdPrtKey=FdPrtKey    
      group By TdBYy, TdBChr, TdBNo, TdRmCd     
        having convert(Decimal(10,3),sum(TdRmWt*(case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1     
       when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1     
       when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1     
       when  TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end))) >0))    
   If @Row > 0    
     begin     
  	Select @ErrMsg= 'Cannt Move the bag to Fg, All the Raw materials Specified In The Order Are Not Used'    
   	GoTo ErrHndlr    
    end 

	Select  @Row= Count('x') from Inserted where FdPrtKey=@wCurrPrtn and exists 
		(Select OrLotNo from OrdRm     
			Join Bag on BPrtKey=OrPrtKey and BOdIdNo=OrOdIdNo and BCoCd=OrCoCd and BOdTc=OrTc and BOdYy=OrYy    
			and BOdChr=OrChr and BOdNo=OrNo and BOdSr=OrSr    
		where BOdTc='SO' and (select HOrLotYN from Head where HCoCd=FdCoCd and HCd=@wHCd)='Y' and OrRmCtg ='D' and   
			(select vPValue3 from vParam where vPTyp='CHR' and vPCoCd=OrCoCd and vPMCd=OrTc and vPSCd=OrChr)='N'  and    
			BPrtKey=FdPrtKey and BIdNo=FdBIdNo and BCoCd= FdCocd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo and 
			OrLotNo <> '' and not exists (
				Select 'a' From Txnd      
                Where TdLotNo <> '' and TdRmCd = OrRmCd and TdLotNo=OrLotNo and TdBIdNo=FdBIdNo and TdCoCd= FdCoCd and TdBYy= FdByy    
					and TdBChr= FdBChr and TdBNo= FdBNo and TdPrtKey=FdPrtKey    
				group By TdBYy, TdBChr, TdBNo, TdRmCd, TdLotNo having convert(Decimal(10,3),sum(TdRmWt*(case when TdFrRmLoc= 'W' and TdFrRmDC= 'C' then -1     
       when TdFrRmLoc= 'W' and TdFrRmDC= 'D' then 1     
       when TdToRmLoc= 'W' and TdToRmDC= 'C' then -1     
       when  TdToRmLoc= 'W' and TdToRmDC= 'D' then 1 else 0 end))) >0) )     	
    
      If @Row > 0    
      begin     

	Select @ErrMsg= 'Cannt Move the Bag to Fg, All the Rm Lot Specified In The Order Are Not Used'    
	 GoTo ErrHndlr    
    End      
  End    

  if @mode= 'A'    
  Begin    
   select @Row= Count('x') from inserted    
    where FdPRtKey=@wCurrPrtn and exists(select 'x' from Txnd     
     where TdPRtKey=FdPrtKey and TdBIdNo=FdBIdNo and TdCoCd= FdCoCd     
     and TdBYy= FdBYy and TdBChr= FdBChr and TdBNo= FdBNo     
     and TdSrNo= 0 and TdYyKey= FdBYyKey and TdAck= 'N')    
   If @Row > 0    
   begin     
    Select @ErrMsg= 'Bag Not Acknowledged In Previous Transaction'    
    GoTo ErrHndlr    
   end    
   --5717
   select @Row= Count('x') from inserted    
    where FdPrtKey=@wCurrPrtn and FdBNo<> 0 and     
     not exists(select 'x' from Bag     
      where FdPrtKey=BPrtKey and FdCoCd= BCoCd     
       and FdPrdOdTc= BOdTc and FdPrdOdYy= BOdYy and FdPrdOdChr= BOdChr     
       and FdPrdOdNo= BOdNo and FdPrdOdSr= BOdSr)    
   If @Row > 0    
   begin     
    Select @ErrMsg= 'Invalid Production Order Number'    
    GoTo ErrHndlr    
   end 

   Select @Row=Count('x') From Inserted Ins Where Ins.FdPrtKey=@wCurrPrtn and FdBNo <> 0 and    
    exists(select 'x' from Bag where FdPrtKey=BPrtKey and FdBIdNo=BIdNo and FdCoCd= BCoCd and FdBYy= BYy    
    and FdBChr= BChr and FdBNo= BNo and BQty= 0)    
   If @Row>0    
   begin    
      Select @ErrMsg= 'Bag with 0 Quantity Cannot be copied to FG'    
      GoTo ErrHndlr    
   end    
  End    
    
  if @mode In ('A','C') /* Manali added mode **/
  Begin    
    /**** FdPlChr  '' or  Valid PMcd('PlChr')  ****/    
    /* Current Partition */    
    If @mode='A' or (@mode='C' and (Update(FdPlChr)))
    Begin
      Select @Row=Count('x') From Inserted Where FdPrtKey=@wCurrPrtn and    
        Not Exists (Select 'x' From Param where PTyp= 'PlChr'    
          and PMCd= FdPlChr and PSCd= '')    
        and FdPlChr<>''    
        If @Row>0    
      begin    
       Select @ErrMsg='Invalid Pl Chr'    
         GoTo ErrHndlr    
      end    
    End
          
    /**** FdPlYy  ''  or  Valid PMcd('Yy') ****/    
    /* Current Partition */    
    If @mode='A' or (@mode='C' and (Update(FdPlYy)))
    Begin
      Select @Row=Count('x') From Inserted Where FdPrtKey=@wCurrPrtn and    
        Not Exists (Select 'x' From Param where PTyp= 'YY'    
          and PMCd= FdPlYy and PScd= '')    
        and FdPlYy<>''    
        If @Row>0    
      begin    
       Select @ErrMsg='Invalid Pl Yy'    
         GoTo ErrHndlr    
      end    
    End
          
    /**** FdPlChr...FdPlNo Either all empty or all filled ****/    
    /* Current Partition */    
    If @mode='A' or (@mode='C' and (Update(FdPlChr) or Update(FdPlYy) or Update(FdPlNo)))
    Begin
      Select @Row=Count('x') From Inserted Where FdPrtKey=@wCurrPrtn and    
        (( (FdPlChr<> '' or FdPlYy<> '') and FdPlNo=0)    
              or ( (FdPlChr= '' or FdPlYy= '') and FdPlNo<>0))        
      If @Row>0    
      begin    
       Select @ErrMsg='Invalid Pl Chr/Yy/No'    
         GoTo ErrHndlr    
      end    
    End
          
    /***** FdInTc  '' or Pvalue1('TcTyp',Pvalue('Tc',FdInTc) = 'In') *******/    
    /* Current Partition */    

    If @mode='A' or (@mode='C' and (Update(FdBYyKey)))
    Begin
      Select @Row=Count('x') From Inserted join Bag on BPrtKey=FdPrtKey and BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo     
       Where FdBYyKey<> BLstYyKey     
        If @Row>0    
      begin    
       Select @ErrMsg= 'FdBYyKey Should Be Equal To The YyKey Of The Last Transaction Of The Bag'    
         GoTo ErrHndlr    
      end    
    End
          
    If @mode='A' or (@mode='C' and (Update(FdSubLoc)))
    Begin
      select @row= count('x') from inserted where FdPrtKey=@wCurrPrtn and    
            not exists (select 'x' from vParam where vPCoCd= FdCoCd and     
         vPTyp= 'FGSUBLOC' and vPMCd= FdSubLoc)    
      if @row> 0      
      begin    
       select @errmsg= 'Finished Goods Sub Location Not Defined'    
       goto ErrHndlr    
      end    
    End
      
  /*******urmi***********************/    
  if @mode= 'C'    
  Begin    
   /* Current Partition */    
    If Update(FdPlYy) or Update(FdPlChr) or Update(FdPlNo) or Update(FdSubLoc)
    Begin    
       select @Row= Count('x') from Deleted del where del.FdPrtKey=@wCurrPrtn and    
        not exists(Select 'x' from inserted ins where ins.FdPrtKey=del.FdPrtKey and ins.FdIdNo=del.FdIdNo     
         and ins.FdPlYy= del.FdPlYy and ins.FdPlChr= del.FdPlChr     
         and ins.FdPlNo= del.FdPlNo 
         and ins.FdSubLoc= del.FdSubLoc)     
        and exists(Select 'x' from Fmd Fmd1 where Fmd1.FmdPrtKey=del.FdPrtKey and Fmd1.FmdCoCd= del.FdCoCd     
         and Fmd1.FmdBYy= del.FdBYy and Fmd1.FmdBChr= del.FdBChr     
         and Fmd1.FmdBNo= del.FdBNo and     
         ((Fmd1.FmdFdYy> del.FdYy)    
          or    
          (Fmd1.FmdFdYy= del.FdYy and Fmd1.FmdFdKey>= del.FdKey)    
         ))    
       If @Row> 0    
       begin        
        Select @ErrMsg= 'Cannot Change As Fg Movement Has Happened For The Bag'    
        GoTo ErrHndlr    
       end    
    End       
   
/* Zubin 212 (Shifted To Nor For Repl Trigger) */     
  End    
  /******** urmi 2.04  **********************/    
    
 End    
 end   
    
	--577chk
       /**** If BNo <> 0 then FdDmCd, FdSfx, FdPrdCmCd should match with Bag *****/    
    /* Current Partition */    
    If @mode='A' or (@mode='C' and (Update(FdDmCd) or Update(FdSfx) or Update(FdDmSz) or Update(FdPrdCmCd)))
    Begin
      select @Row= Count('x') from inserted     
        join Bag on FdPrtKey=BPrtKey and FdBIdNo=BIdNo and FdCoCd= BCoCd and FdBYy= BYy and FdBChr= BChr and FdBNo= BNo     
        join OrdMst on OmPrtKey=BPrtKey and OmIdNo=BOmIdNo and OmCoCd= BCoCd and OmTc=BOdTc and OmYy= BOdYy and OmChr= BOdChr and OmNo= BOdNo     
       where FdPrtKey=@wCurrPrtn and (FdDmCd<> BOdDmCd or FdSfx<> BOdSfx or FdDmSz<> BOdDmSz or FdPrdCmCd<> OmCmCd)    
        and FdBNo<>0    
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Design/Customer Code Fields Do Not Match With Bag'    
       GoTo ErrHndlr    
      end     
    End
	 
 if @mode in ('A', 'C') and @wTcTyp= 'FWB'    
 Begin    
      
  /**** FdRefKey= FdKey and FdRefYy= FdYy of Base Vch ****/    
  /* Current Partition */    
  If @mode='A'  or (@mode='C' and (Update(FdRefYy) or Update(FdRefKey)))
  Begin    
    Select @Row=Count('x') From Inserted Ins Where Ins.FdPrtKey=@wCurrPrtn and    
     Not Exists(Select 'x' from Fgd B where B.FdPRtKEy=Ins.FdPrtKey and B.FdCoCd= Ins.FdCoCd     
      and B.FdYy= Ins.FdRefYy and B.FdKey= Ins.FdRefKey     
      and B.FdBM= 'B' and B.FdBNo<> 0)    
      If @Row>0    
    begin    
     Select @ErrMsg= 'Reference Yy/Key Should Be Equal To Yy/Key Of Base Voucher'    
       GoTo ErrHndlr    
    end    
  End
      
  --       /**** BagNo, FdDmCd, FdSfx, PrdOrd, FdPrdCmCd, Exp Dets should match with Base Vch *****/    
  --/* Current Partition */    
  --If @mode='A'  or (@mode='C' and (Update(FdRefYy) or Update(FdRefKey) Or Update(FdBYy) Or Update(FdBChr) Or 
  --                  Update(FdBNo) Or Update(FdDmCd) or Update(FdSfx) or Update(FdDmSz) or Update(FdQty) or
  --                  Update(FdGrWt) or Update(FdPrdOdTc) or Update(FdPrdOdYy) or Update(FdPrdOdChr) or 
  --                  Update(FdPrdOdNo) or Update(FdPrdOdSr) Or Update(FdPrdCmCd) or
  --                  Update(FdExpCmCd) or Update(FdExpOdTc) or Update(FdExpOdYy) or Update(FdExpOdChr) or 
  --                  Update(FdExpOdNo) or Update(FdExpOdSr) or Update(FdPlYy) or Update(FdPlChr) or Update(FdPlNo) or 
  --                  Update(FdInTc) or Update(FdInYy) or Update(FdInChr) or Update(FdInNo) or Update(FdInExpNo))) 
  --Begin    
  --  Select @Row=Count('x') From Inserted Ins Where Ins.FdPrtKEy=@wCurrPrtn and    
  --   Not Exists(Select 'x' from Fgd B where B.FdPrtKey=Ins.FdPrtKey and B.FdCoCd= Ins.FdCoCd     
  --    and B.FdYy= Ins.FdRefYy and B.FdKey= Ins.FdRefKey and B.FdBYy= Ins.FdBYy and B.FdBChr= Ins.FdBChr     
  --    and B.FdBNo= Ins.FdBNo and B.FdDmCd= Ins.FdDmCd and B.FdSfx= Ins.FdSfx     
  --    and B.FdDmSz= Ins.FdDmSz and B.FdQty= Ins.FdQty and B.FdGrWt= Ins.FdGrWt     
  --    and B.FdPrdOdTc= Ins.FdPrdOdTc and B.FdPrdOdYy= Ins.FdPrdOdYy     
  --    and B.FdPrdOdChr= Ins.FdPrdOdChr and B.FdPrdOdNo= Ins.FdPrdOdNo     
  --    and B.FdPrdOdSr= Ins.FdPrdOdSr and B.FdPrdCmCd= Ins.FdPrdCmCd     
  --    and B.FdExpCmCd= Ins.FdExpCmCd and B.FdExpOdTc= Ins.FdExpOdTc     
  --    and B.FdExpOdYy= Ins.FdExpOdYy and B.FdExpOdChr= Ins.FdExpOdChr     
  --    and B.FdExpOdNo= Ins.FdExpOdNo and B.FdExpOdSr= Ins.FdExpOdSr     
  --    and B.FdPlYy= Ins.FdPlYy and B.FdPlChr= Ins.FdPlChr and B.FdPlNo= Ins.FdPlNo    
  --    and B.FdInTc= Ins.FdInTc and B.FdInYy= Ins.FdInYy and B.FdInChr= Ins.FdInChr     
  --    and B.FdInNo= Ins.FdInNo and B.FdInExpNo= Ins.FdInExpNo)    
  --    If @Row> 0    
  --  begin    
  --   Select @ErrMsg= 'Bag/ Design/ Suffix/ DsgSz/ PrdOrd/ PrdCust/ Exp Dets Should Match With Base Voucher'    
  --     GoTo ErrHndlr    
  --  end    
  --End
      
  /******* FdBYyKey = BLstYyKey ******/     
  /* Current Partition */    
  If @mode='A'  or (@mode='C' and (Update(FdBYyKey)))
  Begin
    Select @Row= Count('x') From Inserted join Bag on BPrtKey=FdPrtKey and BIdNo=FdBIdNo and BCoCd= FdCoCd and BYy= FdBYy and BChr= FdBChr and BNo= FdBNo     
     Where FdPrtKey=@wCurrPrtn and FdBYyKey<> BLstYyKey     
      If @Row> 0    
    begin    
     Select @ErrMsg= 'FdBYyKey Should Be Equal To The YyKey Of The Last Transaction Of The Bag'    
       GoTo ErrHndlr    
    end    
  End
      
 End    
    
    
    
 /*** UPDATES ***/    
    
 /*** Bag forward update, BLoc= FgToBLoc    
      If delete then Reverse BLoc to max record in Txnd ****/    
 if @mode in ('D') and @wTcTyp in ('FB', 'FR')    
 Begin    
 /* Current Partition */    
  select @Row= Count('x') from Deleted Del where Del.FdPrtKey=@wCurrPrtn and Del.FdDespQty<> 0     
  If @Row > 0    
  begin        
   Select @ErrMsg= 'Cannot Delete when DespQty<> 0'    
   GoTo ErrHndlr    
  end    
    
  select @row= count('x') from deleted, Fmd where FdPRtKey=FmdPrtKey and FdCoCd= FmdCoCd and     
    FdYy= FmdFdYy and FdKey= FmdFdKey     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Fg Movement Rec Exists'    
   goto ErrHndlr    
  end    
    
  /********* urmi 2.04 ******************/    
  if @wTcTyp ='FB'     
  Begin    
   /******* Cannot delete If FG Movement has happened For The Bag ******/    
   /* Current Partition */    
   select @Row= Count('x') from Deleted del where del.FdPrtKey=@wCurrPrtn and    
    exists(Select 'x' from Fmd Fmd1 where Fmd1.FmdPrtKey=del.FdPrtKey and Fmd1.FmdCoCd= del.FdCoCd and     
    Fmd1.FmdBYy= del.FdBYy and Fmd1.FmdBChr= del.FdBChr and     
    Fmd1.FmdBNo= del.FdBNo and     
    ( (Fmd1.FmdFdYy> del.FdYy)    
      or    
     (Fmd1.FmdFdYy= del.FdYy and Fmd1.FmdFdKey>= del.FdKey)    
    ) )    
   If @Row> 0    
   begin        
    Select @ErrMsg= 'Cannot Delete As Fg Movement Has Happened For The Bag'    
    GoTo ErrHndlr    
   end    
      
  End    
    
 End     
    
    
Return    
ErrHndlr:    
 BEGIN    
  Select @ErrMsg='|'+@ErrMsg    
  RaisError(@ErrMsg,16,1)    
  Rollback Transaction    
 END    
    
End    


GOGO

ALTER TRIGGER [dbo].[InvDsg_INUPDL] ON [dbo].[InvDsg]
       FOR INSERT,UPDATE,DELETE
AS
BEGIN
	DECLARE @ROW INT
	DECLARE @mode char(1)
	DECLARE @errmsg varchar(255)
	DECLARE @wTcTyp varchar(3)
	DECLARE @wEntTyp varchar(8)
	DECLARE @ReplAgt BIT  
	Declare @wInFob Float
	Declare @wInFobRs Float
	Declare @wInRunningFob Float
	Declare @wInRunningTot Float
	Declare @wInRunningTotRs Float

	DECLARE @wCurrPrtn VarChar(1)
	Declare @wInTCS Float
	Declare @wInTCSRs Float
	Declare @wInCpy char(1)='N'	
	/*** Jen 2.14(LclInv) ***/
	Declare @wInLclOctroiRs Float
	Declare @wInLclVatRs Float
	Declare @wInLclCstRs Float
	/*** Jen 2.14(LclInv) ***/
		
	IF @@ROWCOUNT=0		RETURN
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'InvDsg' )
		RETURN

	select @mode=case (select count('x') from inserted) when 0 then 'D'
		     else
			case (select count('x') from deleted) when 0 then 'A' else 'C' end
		     end	

	/* Sachin 3.02 - Constant For Current Partition */
	Select @wCurrPrtn = 'C'	



	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt= Convert(BIT,SessionProperty('Replication_Agent'))
	
	If  (@ReplAgt<> 1)
	Begin

	/*** Storing Tc Type In The Variable @wTcTyp ***/
	if @mode in ('A','C')
	begin
		select @wTcTyp= max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= IdTc and PSCd= ''
	end

	if @mode in ('D')
	begin
		select @wTcTyp=max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= IdTc and PSCd= '' 
	end
	
	select @wInCpy='Y' where exists (Select (Cpy) from inserted where Cpy='Y')

	/*** Storing Entry Type In The Variable @wEntTyp ***/
	select @wEntTyp= PValue1 from Param  where PTyp= 'TCTYP' and PMCd= @wTcTyp and PSCd= ''
	--5717
	/******** Only One TcTyp records can be Inserted at a time.********/	
	select @Row= Count(distinct(PValue)) from inserted Join param On
		PTyp= 'TC' and PMCd= IdTc and PScd= '' Where IdPrtKey=@wCurrPrtn 
	If @Row> 1
	begin
		Select @ErrMsg= 'Transactions of only One TcType can be Acted On'
	  	GoTo ErrHndlr
	end
	--5717
	/******** Only One TcTyp records can be Deleted at a time. *********/	
	select @Row= Count(distinct(PValue)) from deleted Join param on
		PTyp= 'TC' and PMcd= IdTc and PSCd='' Where IdPrtKey=@wCurrPrtn
	If @Row> 1
	begin
		Select @ErrMsg= 'Transactions of only One TcType can be Acted on'
		GoTo ErrHndlr
	end

	if @mode = 'D'
	Begin 
		/* Current Partition */
		select @Row= Count('x') from deleted join InvHd on InPrtKey=IdPrtKey and InIdNo=IdInIdNo 
			and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo where InCls= 'Y' and IdPrtKey=@wCurrPrtn

		If @Row > 0
		begin 
			Select @ErrMsg= 'Cannot Delete As Invoice Is Closed'
			GoTo ErrHndlr
		end
	End 

	if @wInCpy='N' 
	begin
	if @mode = 'A'
	Begin
		/* Current Partition */
		select @Row= Count('x') from inserted join InvHd on InPrtKey=IdPrtKey and InIdNo=IdInIdNo
			and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo where IdPrtKey=@wCurrPrtn and InCls= 'Y' 
		If @Row> 0 
		begin 
			Select @ErrMsg= 'Cannot Add As Invoice Is Closed'
			GoTo ErrHndlr
		end
	End

	if @mode = 'C'
	Begin 
		select @Row= Count('x') from deleted del 
			join inserted ins on ins.IdPrtKey=del.IdPrtKey and ins.IdIdNo=del.IdIdNo and ins.IdCoCd= del.IdCoCd and ins.IdTc= del.IdTc and ins.IdYy= del.IdYy 
				and ins.IdChr= del.IdChr and ins.IdNo= del.IdNo and ins.IdSr= del.IdSr 
			join InvHd on InPrtKey=del.IdPrtKey and InIdNo=del.IdInIdNo and InCoCd= del.IdCoCd and InTc= del.IdTc and InYy= del.IdYy 
				and InChr= del.IdChr and InNo= del.IdNo 
			where InCls= 'Y' 
		If @Row> 0
		begin 
			Select @ErrMsg= 'Cannot Change As Invoice Is Closed'
			GoTo ErrHndlr
		end
	End 
	--/*** (Jen 2.13) 28/11/06 ***/


    
    
		/******** Cannot Change IdGldLs if record exists in InvZRm for the Invoice ********/ 	
		/* Current Partition */
	If @mode='C' and update(IdGldLs) 
    Begin		    
		  select @Row= Count('x') from deleted del where del.IdPrtKey=@wCurrPrtn and
			  not exists (Select 'x' from inserted ins where ins.IdPrtKey=del.IdPrtKey and ins.IdIdNo=del.IdIdNo
					  and ins.IdCoCd= del.IdCoCd and ins.IdTc= del.IdTc and ins.IdYy= del.IdYy and ins.IdChr= del.IdChr 
					  and ins.IdNo= del.IdNo and ins.IdSr= del.IdSr and Round(ins.IdGldLs, 3)= Round(del.IdGldLs, 3)) 
			  and exists(Select * from InvZRm where IzCoCd= del.IdCoCd and 
					  IzTc= del.IdTc and IzYy= del.IdYy and IzChr= del.IdChr and IzNo= del.IdNo) 
		  If @Row> 0
		  begin 
			  Select @ErrMsg= 'Cannot Change Gold Loss As Custom Matching Has Been Done'
			  GoTo ErrHndlr
		  end
    End
    
	If @mode='C' and update(IdDmCd) or Update(IdSfx) or Update(IdDmSz) or Update(IdQty) 
    Begin		    
		  select @row= count('x') from inserted ins where 
			  not exists (select 'x' from deleted del where 
					  ins.IdIdNo=del.IdIdNo  and ins.IdQty>= del.IdQty 
					  and ins.IdDmCd= del.IdDmCd and ins.IdSfx= del.IdSfx 
					  and ins.IdDmSz= del.IdDmSz) 
			  and exists(Select 'x' from RetMemDet where RedCoCd= ins.IdCoCd and RedIdTc= ins.IdTc 
					  and RedIdYy= ins.IdYy and RedIdChr= ins.IdChr and RedIdNo= ins.IdNo 
					  and RedIdSr= ins.IdSr) 
		  If @Row> 0
		  begin 
			  Select @ErrMsg= 'Cannot Change Design, Suffix, Size, Qty As Memo Return Detail Exists'
			  GoTo ErrHndlr
		  end
    End
	--5717
	if @mode = 'C'
 	Begin
		/******** Key Cannot Be Changed ********/ 
		/* Current Partition */
		If update(IdPRtKey) or Update(IdIdNo) or Update(IdCoCd) or Update(IdTc) or Update(IdYy) or Update(IdChr)
		    or Update(IdNo) or Update(IdSr)
    Begin		    
		  select @Row=Count('x') from deleted del where del.IdPrtKey=@wCurrPrtn and
			   not exists (Select 'x' from inserted ins where ins.IdPRtKey=del.IdPrtKey and ins.IdIdNo=del.IdIdNo
					  and ins.IdCoCd= del.IdCoCd and ins.IdTc= del.IdTc and ins.IdYy= del.IdYy and ins.IdChr= del.IdChr 
					  and ins.IdNo= del.IdNo and ins.IdSr= del.IdSr)
		  If @Row > 0
		  begin 
			  Select @ErrMsg='Cannot Change Company Code/ Tc/Yy/Chr/No/Sr'
			  GoTo ErrHndlr
		  end
    End
	--5717
	If update(IdDmCd) or Update(IdExpCmCd) or Update(IdExpOdTc) or Update(IdExpOdYy) or Update(IdExpOdChr)
		    or Update(IdExpOdNo) or Update(IdExpOdSr)
    Begin		    
		  select @Row= Count('x') from deleted del where del.IdPrtKey=@wCurrPrtn and
			   not exists (Select 'x' from inserted ins where ins.IdPrtKey=del.IdPrtKey and ins.IdIdNo=del.IdIdNo
					  and ins.IdCoCd= del.IdCoCd and ins.IdTc= del.IdTc and ins.IdYy= del.IdYy and ins.IdChr= del.IdChr 
					  and ins.IdNo= del.IdNo and ins.IdSr= del.IdSr and ins.IdDmCd= del.IdDmCd and ins.IdExpCmCd= del.IdExpCmCd 
					  and ins.IdExpOdTc= del.IdExpOdTc and ins.IdExpOdYy= del.IdExpOdYy and ins.IdExpOdChr= del.IdExpOdChr 
					  and ins.IdExpOdNo= del.IdExpOdNo and ins.IdExpOdSr= del.IdExpOdSr)
		  If @Row> 0
		  begin 
			  Select @ErrMsg= 'Cannot Change Design Code/ Export Order'
			  GoTo ErrHndlr
		  end
    End
	End	

	if @mode in ('A')
	Begin
		select @row= count('x') from inserted where @wTcTyp= 'IN' and 
			not exists (select 'x' from OrdDsg where OdCoCd= IdCoCd 
			and OdTc= IdExpOdTc and OdYy= IdExpOdYy and OdChr= IdExpOdChr 
			and OdNo= IdExpOdNo and OdSr= IdExpOdSr and (OdCls= 'N' or OdHld='N'))
		if @row> 0 
		begin
			select @errmsg= 'Export Order Not Defined/Order Closed/Order Design is on Hold'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where IdPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and 
			not exists (select 'x' from OrdDsg where OdPrtKey=IdPrtKey and OdCoCd= IdCoCd 
			and OdTc= IdExpOdTc and OdYy= IdExpOdYy and OdChr= IdExpOdChr 
			and OdNo= IdExpOdNo and OdSr= IdExpOdSr and OdHld= 'N') 
		if @row> 0 
		begin
			select @errmsg= 'Order Design is on Hold'
			goto ErrHndlr
		end
		/**************************Geeta*************************/
		/*** Jay 3.3(FG) ***/
		if @wTcTyp= 'FG' 
		begin
			select @errmsg= 'Cannot Add Record For ''FOT'' TcTyp'
			goto ErrHndlr
		end	
	End

	if @mode in ('A','C')
	begin
		/* Current Partition */
		If @mode='A' or (@mode='C' and (update(IdPrtKey) or Update(IdCoCd) or Update(IdTc) or Update(IdYy)
		                              or Update(IdChr) or Update(IdNo)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and not exists (select 'x' from InvHd where 
				  InPrtKey=IdPrtKey and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and 
				  InChr= IdChr and InNo= IdNo)
		  if @row>0 
		  Begin
			  select @errmsg='Voucher Not Defined In Invoice Header'
			  goto ErrHndlr
		  end
    End
    --5717
	If @mode='A' or (@mode='C' and (Update(IdSr)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and IdSr<=0 
		  if @row>0 
		  Begin
			  select @errmsg= 'Invoice Design Serial Number Should Be > Zero'
			  goto ErrHndlr
		  end
    End
	--5717
	If @mode='A' or (@mode='C' and (update(IdExpCmCd)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and not exists (select 'x' from OrdMst where 
				  OmPrtKey=IdPrtKey and OmCoCd= IdCoCd and OmTc= IdExpOdTc and OmYy= IdExpOdYy and 
				  OmChr= IdExpOdChr and OmNo= IdExpOdNo and OmCmCd= IdExpCmCd)
		  if @row>0 
		  Begin
			  select @errmsg= 'Export Customer Code Not Defined'
			  goto ErrHndlr
		  end
    End
	--5717
	If @mode='A' or (@mode='C' and (update(IdQty)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and
			  IdQty<  0
		  if @row>0 
		  Begin
			  select @errmsg= 'Quantity Should Be >= Zero'
			  goto ErrHndlr
		  end
    End
	--5717
	If @mode='A' or (@mode='C' and (update(IdKt)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and 
			  not exists (select 'x' from Param where 
				      PTyp= 'KT' and PMCd= IdKt)
		  if @row>0 
		  Begin
			  select @errmsg= 'Karatage Not Defined'
			  goto ErrHndlr
		  end
    End
	--5717
	If @mode='A' or (@mode='C' and (update(IdExpCmCd)))
    Begin		    
		  select @row= count('x') from inserted 
				  join CustMst Ord on CmCtg= 'C' and CmCd= IdExpCmCd 
				  where IdPrtKey=@wCurrPrtn and Ord.CmCurCd <> 
					  (select Inv.CmCurCd from InvHd 
					  join CustMst Inv on Inv.CmCtg= (case when @wTcTyp= 'IN' then 'C' 
						  when @wTcTyp= 'CTB' then 'T' else '' end) and Inv.CmCd= InCmCd 
					  where InPrtKey=IdPrtKey and InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr 
						  and InNo= IdNo 
					  ) 
		  if @row> 0 
		  begin
			  select @errmsg= 'Currency Of The Export Order And Invoice Not Matching'
			  goto ErrHndlr
		  end		  
    End
    
		If @mode='A' or (@mode='C' and (update(IdDmCd) or Update(IdSfx) or Update(IdDmSz)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and not exists (select 'x' from OrdDsg where 
				  OdPRtKey=IdPrtKey and OdCoCd= IdCoCd and OdTc= IdExpOdTc and OdYy= IdExpOdYy and 
				  OdChr= IdExpOdChr and OdNo= IdExpOdNo and OdSr= IdExpOdSr and OdDmCd= IdDmCd and 
				  OdDmSz= IdDmSz and OdSfx= IdSfx and OdDmTcTyp='DM' )
		  if @row>0 
		  Begin
			  select @errmsg= 'Design/ Suffix/ Size Does Not Belong To The Export Order'
			  goto ErrHndlr
		  end
    End
    
    
	If @mode='A' or (@mode='C' and (update(IdVaCtg)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and
			  not exists (select 'x' from Param where 
				      PTyp= 'VACTG' and PMCd= IdVaCtg)
		  if @row>0 
		  Begin
			  select @errmsg= 'Value Addition Category Not Defined'			  goto ErrHndlr
		  end
    End
    
		/* Current Partition */
		If @mode='A' or (@mode='C' and (update(IdGldLs)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and IdGldLs< 0	
		  if @row>0 
		  Begin
			  select @errmsg= 'Gold Loss Should Be >= Zero'			  goto ErrHndlr
		  end
    End
    
    
		/* Current Partition */
		If @mode='A' or (@mode='C' and (update(IdFixPrc)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and
			  not exists (select 'x' from Param where 
				      PTyp= 'YN' and PMCd= IdFixPrc)
		  if @row>0 
		  Begin			  select @errmsg= 'Enter Fix Price (Y/N)'
			  goto ErrHndlr
		  end
    End
    
		If @mode='A' or (@mode='C' and (update(IdGldAs)))
    Begin		    
		  select @row= count('x') from inserted where IdPrtKey=@wCurrPrtn and IdGldAs<> '' and 
			  not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= IdGldAs)
		  if @row> 0 
		  Begin
			  select @errmsg= 'Gold As Value Not Defined'
			  goto ErrHndlr
		  end
    End
    
		-- ****** Sachin 2.13 - 26-05-2006 ******
		/* Current Partition */
		If @mode='A' or (@mode='C' and (update(IdLabAs)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and	
			  not exists (select 'x' from Param where 
				      PTyp= 'GLDAS' and PMCd= IdLabAs)
		  if @row>0 
		  Begin
			  select @errmsg= 'Labour As Value Not Defined'
			  goto ErrHndlr
		  end
    End
    
		/* Current Partition */
		If @mode='A' or (@mode='C' and (update(IdHZValEqInv)))
    Begin		    
		  select @row=count('x') from inserted where IdPrtKey=@wCurrPrtn and
			  not exists (select 'x' from Param where 
				      PTyp= 'YN' and PMCd= IdHZValEqInv)
		  if @row>0 
		  Begin			  select @errmsg= 'Specify Whether Customs Value Equal To Invoice Value (Y/N)'
			  goto ErrHndlr
		  end
    End
	end
	end

	if @mode = 'D' 	
	Begin
		select @row= count('x') from deleted del join RetMemDet on RedCoCd= del.IdCoCd 
					and RedIdTc= del.IdTc and RedIdYy= del.IdYy and RedIdChr= del.IdChr 
					and RedIdNo= del.IdNo and RedIdSr= del.IdSr 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Return Memo Detail Exists'
			goto ErrHndlr
		end
		   Delete InvRm From deleted, InvRm where IrPRtKey=IdPrtKey and IrIdIdNo=IdIdNo and IrCoCd= IdCoCd and IrTc= IdTc and IrYy= IdYy and 
					IrChr= IdChr and IrNo= IdNo and IrSr= IdSr  
   
		   Delete InvLab From deleted, InvLab where IlPrtKey=IdPrtKey and IlIdIdNo=IdIdNo and iLCoCd= IdCoCd and iLTc= IdTc and iLYy= IdYy and 
					iLChr= IdChr and iLNo= IdNo and iLSr= IdSr


		Update InvFgd
		Set IfSr=0
		From Deleted Where 
		IfCoCd= IdCoCd and IfTc= IdTc and IfYy= IdYy and 
					IfChr= IdChr and IfNo= IdNo and IfSr= IdSr 
	End
	
		if @mode in('C', 'D')		
		Begin
	
		  If @mode='D' or (@mode='C' and (update(IdiFOB)))
      Begin  
			  /* sv.44 getting TCS % from head and calculating tcs amount and added in totals  **/
			  update InvHd set 
			  @wInFob= convert (decimal (16, 2), InFOB- IsNull((Select sum(del.IdiFOB) from deleted del 
			  where del.IdCoCd= InCoCd and del.IdTc= InTc and del.IdYy= InYy 
				  and del.IdChr= InChr and del.IdNo= InNo), 0)),
			  @wInFobRs= convert (decimal (16, 0), (@wInFob)
					  * (case when InRsMulDiv= 'M' then inCnvRt else 1 end) 
					  / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) ), 
			  InFOB= @wInFob, 
			  InFOBRs= (case when @wTcTyp= 'IN' then @wInFobRs else 0 end), 
			  @wInLclOctroiRs= convert (decimal (16, 2), (Case when @wTcTyp= 'IN' and InLclInvYn = 'Y' then 
								  (@wInFob * (InCmOctroiPer / 100)) else 0 end)),
			  @wInLclVatRs= convert (decimal (16, 2), (Case when @wTcTyp= 'IN' and InLclInvYn = 'Y' then 
								  ((@wInFOB * (1 + (InCmOctroiPer/100 ))+InLclOthrValRs)* (InHVatPer /100 )) else 0 end)),
			  @wInLclCstRs= convert (decimal (16, 2), (Case when @wTcTyp= 'IN' and InLclInvYn = 'Y' then 
								  ((@wInFOB * (1 + (InCmOctroiPer/100 ))+InLclOthrValRs)* (InHCstPer /100 )) else 0 end)),
			  InLclOctroiRs= @wInLclOctroiRs, 
			  InLclVatRs= @wInLclVatRs, 
			  InLclCstRs= @wInLclCstRs,
		      @wInRunningTot= convert (decimal (16, 2), (Case InLclInvYn 
					    when 'N' then @wInFOB +(case InTotAs when 'FOB' then (case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)
											    when 'CIF' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)
											    when 'C&F' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)
											    when 'C&I' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)
											    else 0 end)
					    when 'Y' then  @wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InLclOthrValRs 
					    else 0 end)),
		    @wInRunningTotRs= (Case InLclInvYn 
					    when 'N' then @wInFOBRs  +(case InTotAs when 'FOB' then (case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)
												    when 'CIF' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)
												    when 'C&F' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)
												    when 'C&I' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)
												    else 0 end)
					    when 'Y' then convert (decimal (16, 0), 
						    (@wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InLclOthrValRs) 
						    * (case when InRsMulDiv= 'M' then inCnvRt else 1 end) 
						    / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) )
					    else 0 end)	,				
			   			  
			  @wInTCS = convert (decimal (16, 2), (@wInRunningTot * InTCSPer) /100),
			  @wInTCSRs = convert (decimal (16, 2), @wInTCS * (case when InRsMulDiv= 'M' then inCnvRt else 1 end) 
			  		                                        / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) ), 
			  InTCS	= @wInTCS,
			  InTCSRs = @wInTCSRs,			
		    InTot= convert (decimal (16, 2), (Case InLclInvYn 
					    when 'N' then @wInRunningTot + @wInTCS 
					    when 'Y' then @wInRunningTot + @wInTcs
					    else 0 end)),
		    InTotRs= (Case InLclInvYn 
					  when 'N' THEN @wInRunningTotRs + @wInTCSRs
					  when 'Y' then 
					  convert (decimal (16, 0), @wInRunningTotRs + @wInTCSRs) else 0 end)
			  from deleted where InIdNo=IdInIdNo and InCoCd= IdCoCd and InTc= IdTc and InYy= IdYy and InChr= IdChr and InNo= IdNo 
      End
    	 If @mode='D' or (@mode='C' and (update(IdQty)))
     Begin  
			  update OrdDsg set
			  Modusr='MIC', OdExpQty= OdExpQty- IsNull((Select sum(del.IdQty) from deleted del where 
				   @wTcTyp= 'IN' and del.IdCoCd= OdCoCd and del.IdExpOdTc= OdTc 
				  and del.IdExpOdYy= OdYy and del.IdExpOdChr= OdChr 
				  and del.IdExpOdNo= OdNo and del.IdExpOdSr= OdSr), 0)
			  from deleted where @wTcTyp= 'IN' and OdCoCd= IdCoCd and OdTc= IdExpOdTc 
				  and OdYy= IdExpOdYy and OdChr= IdExpOdChr and OdNo= IdExpOdNo 
				  and OdSr= IdExpOdSr
  	
			  /*** 2.12 (Jen) ***/
			  /* Current Partition */
			  update OrdDsg set Modusr='MIC', OdCls = (case when OdOrdQty= OdExpQty and Round(OdOrdQty, 1)<> 0 then 'Y' else 'N' end) 
			  from deleted del where OdPRtKEy=IdPRtKey and OdCoCd= IdCoCd and OdTc= IdExpOdTc 
				  and OdYy= IdExpOdYy and OdChr= IdExpOdChr and OdNo= IdExpOdNo 
				  and OdSr= IdExpOdSr
			  /*** 2.12 (Jen) ***/
		  End
		End
	
		if @mode in('C', 'A')		
		Begin		
		 If @mode='A' or (@mode='C' and (update(IdiFOB)))
     Begin  			
		/* sv.44 getting TCS % from head and calculating tcs amount and added in totals  **/
			update InvHd set
			@wInFob = convert (decimal (16, 2), InFOB+ IsNull((Select sum(ins.IdiFOB) from inserted ins where 
				ins.IdPrtKey=InPrtKey and ins.IdCoCd= InCoCd and ins.IdTc= InTc and ins.IdYy= InYy 
				and ins.IdChr= InChr and ins.IdNo= InNo), 0)), 
			@wInFobRs= convert (decimal (16, 0), (@wInFob)
					* (case when InRsMulDiv= 'M' then inCnvRt else 1 end) 
					/ (case when InRsMulDiv= 'D' then inCnvRt else 1 end) ), 
			InFOB= @wInFob, 
			InFOBRs= (case when @wTcTyp= 'IN' then @wInFobRs else 0 end), 
			@wInLclOctroiRs= convert (decimal (16, 2), (Case when @wTcTyp= 'IN' and InLclInvYn = 'Y' then 
								(@wInFob * (InCmOctroiPer / 100)) else 0 end)),
			@wInLclVatRs= convert (decimal (16, 2), (Case when @wTcTyp= 'IN' and InLclInvYn = 'Y' then 
								((@wInFOB * (1 + (InCmOctroiPer/100 ))+InLclOthrValRs)* (InHVatPer /100 )) else 0 end)),
			@wInLclCstRs= convert (decimal (16, 2), (Case when @wTcTyp= 'IN' and InLclInvYn = 'Y' then 
								((@wInFOB * (1 + (InCmOctroiPer/100 ))+InLclOthrValRs)* (InHCstPer /100 )) else 0 end)),
			InLclOctroiRs= @wInLclOctroiRs, 
			InLclVatRs= @wInLclVatRs, 
			InLclCstRs= @wInLclCstRs, 			

			@wInRunningTot= convert (decimal (16, 2), (Case InLclInvYn 
					    when 'N' then @wInFOB +(case InTotAs when 'FOB' then (case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)
											    when 'CIF' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)
											    when 'C&F' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)
											    when 'C&I' then (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End)
											    else 0 end)
					    when 'Y' then  @wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InLclOthrValRs 
					    else 0 end)),
		    @wInRunningTotRs= (Case InLclInvYn 
					    when 'N' then @wInFOBRs  +(case InTotAs when 'FOB' then (case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)
												    when 'CIF' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)
												    when 'C&F' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)
												    when 'C&I' then (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End)
												    else 0 end)
					    when 'Y' then convert (decimal (16, 0), 
						    (@wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InLclOthrValRs)
						    * (case when InRsMulDiv= 'M' then inCnvRt else 1 end) 
						    / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) )
					    else 0 end)	,				
			   			  
			  @wInTCS = convert (decimal (16, 2), (@wInRunningTot * InTCSPer) /100),
			  @wInTCSRs = convert (decimal (16, 2), @wInTCS * (case when InRsMulDiv= 'M' then inCnvRt else 1 end) 
			  		                                        / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) ), 
			  InTCS	= @wInTCS,
			  InTCSRs = @wInTCSRs,			
		    InTot= convert (decimal (16, 2), (Case InLclInvYn 
					    when 'N' then @wInRunningTot + @wInTCS 
					    when 'Y' then @wInRunningTot + @wInTcs
					    else 0 end)),
		    InTotRs= (Case InLclInvYn 
					  when 'N' THEN @wInRunningTotRs + @wInTCSRs
					  when 'Y' then 
					  convert (decimal (16, 0), @wInRunningTotRs + @wInTCSRs) else 0 end)
			
			from inserted where InPrtKey=IdPrtKey and InIdNo=IdInIdNo 
			End
		 If @mode='A' or (@mode='C' and (update(IdQty)))
     Begin  
			  update OrdDsg set ModUsr='MIC',
			  OdExpQty= OdExpQty+ IsNull((Select sum(ins.IdQty) from inserted ins where 
				  @wTcTyp= 'IN' and ins.IdPrtKey=OdPrtKey and ins.IdCoCd= OdCoCd and ins.IdExpOdTc= OdTc 
				  and ins.IdExpOdYy= OdYy and ins.IdExpOdChr= OdChr 
				  and ins.IdExpOdNo= OdNo and ins.IdExpOdSr= OdSr), 0)
			  from inserted where @wTcTyp='IN' and IdPrtKey=@wCurrPrtn and OdCoCd= IdCoCd and OdTc= IdExpOdTc 
				  and OdYy= IdExpOdYy and OdChr= IdExpOdChr and OdNo= IdExpOdNo 
				  and OdSr= IdExpOdSr
  	
			  update OrdDsg set ModUsr='MIC',OdCls = (case when OdOrdQty= OdExpQty and Round(OdOrdQty, 1)<> 0 then 'Y' else 'N' end) 
			  from inserted ins where IdPrtKey=@wCurrPrtn and OdPRtKey=IdPrtkey and OdIdNo=IdOdIdNo and OdCoCd= IdCoCd and OdTc= IdExpOdTc 
				  and OdYy= IdExpOdYy and OdChr= IdExpOdChr and OdNo= IdExpOdNo 
				  and OdSr= IdExpOdSr
		  End	
		End
	End	
return 
ErrHndlr: 
begin 
	select @errmsg='|'+@errmsg 
	raiserror(@errmsg,16,1) 
	rollback transaction 
end 
END 


GOGO

ALTER TRIGGER [dbo].[InvFgd_INUPDL] ON [dbo].[InvFgd]  
  FOR INSERT,UPDATE,DELETE   
  NOT FOR REPLICATION
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
 Declare @wInCpy char(1)='N' 
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
--  select @wInCpy ='Y' where exists (Select 'a' from Inserted where inserted.Cpy='Y') 

 if @mode in ('D')  
 begin  
  select @wTcTyp= max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= IfTc and PSCd= ''  
 end  
   
 /*** Storing Entry Type In The Variable @wEntTyp ***/  
 select @wEntTyp= PValue1 from Param  where PTyp= 'TCTYP' and PMCd= @wTcTyp and PSCd= ''  
 Select @wCurrPrtn='C'  
  
 /* **** Checking Whether a Normal Transaction or Replicated One **** */  
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))  
   
  
   
   
  if @mode = 'D'    
  Begin  
   update Fgd set   
   FdDespQty= FdDespQty- IsNull((Select sum(del.IfQty) from deleted del where   
    @wTcTyp= 'IN' and del.IfPrtKey=FdPrtKey and del.IfFdIdNo=FdIdNo and del.IfCoCd= FdCoCd and del.IfFdTc= FdTc   
    and del.IfFdYy= FdYy and del.IfFdChr= FdChr and del.IfFdNo= FdNo   
    and del.IfFdSr= FdSr), 0)   
   from deleted where IfPrtkey=@wCurrPrtn and @wTcTyp= 'IN' 
		and FdIdNo=IfFdIdNo 
   end
   
	


if @Mode in ('A','C')
begin
update Bag set   
	   BCls= (Case when ins.IfSr>0 Then 'Y' Else 'N' End),   
	   BClsDt= (Case when ins.IfSr>0 Then InDt Else '01/01/80' End)
From inserted ins Join Bag b On b.BIdNo=ins.IfBIdNo and b.BCoCd= ins.IfCoCd and b.BYy= ins.IfFdBYy   and 
						b.BChr= ins.IfFdBChr and b.BNo= ins.IfFdBNo and ins.IfPrtKey=@wCurrPrtn 
Join InvHd ih on InPrtKey=ins.IfPrtKey and InIdNo=ins.IfInIdNo and InCoCd= ins.IfCoCd And InTc = ins.IfTc   
	                     And InYy = ins.IfYy And InChr= ins.IfChr And InNo= ins.IfNo
end

if @Mode in ('D')
begin
 update Bag set   
	   BCls= 'N',   
	   BClsDt= '01/01/80'
  From deleted del Join Bag b On b.BIdNo=del.IfBIdNo and b.BCoCd= del.IfCoCd and b.BYy= del.IfFdBYy   and 
						b.BChr= del.IfFdBChr and b.BNo= del.IfFdBNo and del.IfPrtKey=@wCurrPrtn 
   end

  if @mode in ('A')  
  Begin  
   update Fgd set   
   FdDespQty= FdDespQty+ IsNull((Select sum(ins.IfQty) from inserted ins where   
    @wTcTyp= 'IN' and ins.IfFdIdNo=FdIdNo and ins.IfCoCd= FdCoCd and ins.IfFdTc= FdTc   
    and ins.IfFdYy= FdYy and ins.IfFdChr= FdChr and ins.IfFdNo= FdNo and ins.IfFdSr= FdSr), 0)   
   from inserted where @wTcTyp= 'IN' and FdIdNo=IfFdIdNo  
    and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr  
   
   
 End  
 --5717
 /******** Only One TcTyp records can be Inserted at a time.********/   
 select @Row= Count(distinct(PValue)) from inserted Join param On  
  PTyp= 'TC' and PMCd= IfTc and PSCd= '' Where IfPrtKey=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions Of Only One TcType Can Be Acted On'  
    GoTo ErrHndlr  
 end  
 --5717 
 /******** Only One TcTyp records can be Deleted at a time. *********/   
 select @Row= Count(distinct(PValue)) from deleted Join param on  
  PTyp= 'TC' and PMcd= IfTc and PSCd= '' Where IfPrtKEy=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions Of Only One TcType Can Be Acted On'  
  GoTo ErrHndlr  
 end  
  
  /******** Cannot Add/ Delete/ Change If Invoice Is Closed ********/   
 /* Current Partition */  
if @wInCpy ='N' 
begin
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
  
 -- /* Current Partition */  
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
  select @Row=Count('x') from deleted del where del.IfPrtKey=@wCurrPrtn and  
    not exists (Select 'x' from inserted ins where ins.IfPrtKey=del.IfPrtKey and ins.IfIdNo=del.IfIdNo   
     and ins.IfCoCd= del.IfCoCd and ins.IfTc= del.IfTc and ins.IfYy= del.IfYy and ins.IfChr= del.IfChr   
     and ins.IfNo= del.IfNo and ins.IfSrNo= del.IfSrNo)  
  If @Row > 0  
  begin   
   Select @ErrMsg='Cannot Change Company Code/ Tc/Yy/Chr/No/SrNo'  
   GoTo ErrHndlr  
  end  
  --5717
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
  --5717
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


  
 if @mode= 'A'  
 Begin  
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfSr= 0 and   
    exists (select 'x' from InvDsg where IdPrtKey=IfPrtKey and IdInIdNo=IfInIdNo and IdCoCd= IfCoCd   
    and IdTc= IfTc and IdYy= IfYy and IdChr= IfChr and IdNo= IfNo)  
  if @row>0   
  begin  
   select @errmsg= 'Cannot Enter Record With Sr = 0 (i.e. When Inv Design Record Exists)'  
   goto ErrHndlr  
  end  
  
  select @row= count('x') from inserted where 
   (exists(Select 'x' from InvRm where IrPrtKey=IfPrtKey  
     and IrCoCd= IfCoCd and IrTc= IfTc and IrYy= IfYy and IrChr= IfChr and IrNo= IfNo and IrSr= IfSr) or  
   exists(Select 'x' from InvLab where IlPrtKey=IfPRtKey and iLCoCd= IfCoCd and iLTc= IfTc and iLYy= IfYy and   
     iLChr= IfChr and iLNo= IfNo and iLSr= IfSr))   
  if @row> 0   
  begin  
   select @errmsg= 'Cannot Add As Invoice Raw Material/ Labour Record Exists'  
   goto ErrHndlr  
  end  
  
  select @row= count('x') from inserted ins join Fgd Fd1 on   
	FdIdNo=IfFdIdNo 
   where @wTcTyp= 'IN' and IfSr= 0 and 
    Round(IsNull((select OdOrdQty- OdExpQty from OrdDsg where 
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
 End  
  
 if @mode in ('A','C')  
 Begin  
  /****** Voucher (Tc/Yy/Chr/No) should exist in InvHd ******/  
  --577chk
  select @row= count('x') from inserted where IfPRtKey=@wCurrPrtn   
   and not exists (select 'x' from InvHd where InPrtKey=IfPrtKey and InIdNo=IfInIdNo  
    and InCoCd= IfCoCd and InTc= IfTc and InYy= IfYy and InChr= IfChr and InNo= IfNo)  
  if @row> 0   
  begin  
   select @errmsg= 'Voucher Not Defined In Invoice Header'  
   goto ErrHndlr  
  end  
  
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfSr<> 0   
    and not exists (select 'x' from InvDsg where   
         IdPrtKey=IfPrtKey and IdCoCd= IfCoCd and IdTc= IfTc and IdYy= IfYy   
         and IdChr= IfChr and IdNo= IfNo and IdSr= IfSr)  
  if @row>0   
  begin  
   select @errmsg= 'Voucher Sr Not Defined In Invoice Design'  
   goto ErrHndlr  
  end  
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfQty<= 0   
  if @row>0   
  begin    select @errmsg= 'Inv Bag Qty Should Be > Zero'  
   goto ErrHndlr  
  end  
  select @row= count('x') from inserted where not exists (select 'x' from Fgd where   
    FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and   
    FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and   
    FdBYy= IfFdBYy and FdBChr= IfFdBChr and FdBNo= IfFdBNo)   
  if @row>0   
  begin  
   select @errmsg= 'Finished Goods Voucher/ Bag Number Not Defined'  
   goto ErrHndlr  
  end  
  --5717
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfSr< 0   
  if @row>0   
  begin  
   select @errmsg='Invoice Bag Serial Cannot Be < Zero'  
   goto ErrHndlr  
  end  
  --5717
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and IfSrNo<= 0   
  if @row>0   
  begin  
   select @errmsg='Invoice Bag Serial Number Cannot Be <= Zero'  
   goto ErrHndlr  
  end 
  --5717
  select @row= count('x') from inserted where IfPrtKey=@wCurrPrtn and  
   not exists (select 'x' from Param where PTyp='TC' and PMCd= IfFdTc and   
        PValue in ('FB', 'FR')    )  
  if @row>0   
  begin  
   select @errmsg= 'Fg Voucher Tc Not Defined'  
   goto ErrHndlr  
  end  
  --5717
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
  --5717
  select @row=count('x') from inserted where IfPrtKey=@wCurrPrtn and not exists (select 'x' from Fgd where   
    FdPRtKey=IfPrtKey and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and   
    FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and FdBIdNo=IfBIdNo and FdBYy= IfFdBYy   
    and FdBChr= IfFdBChr and FdBNo= IfFdBNo and FdQty= IfQty and FdGrWt= IfGrsWt)   
  if @row>0   
  begin  
   select @errmsg= 'Qty/ Weight Should Be The Actual Qty/ Weight In The Bag'  
   goto ErrHndlr  
  end  


	/* doubt */  
  --select @row= count('x') from inserted   
  -- join InvDsg on IdPrtKey=IfPrtKey and IdCoCd= IfCoCd and IdTc= IfTc and IdYy= IfYy and IdChr= IfChr   
  --  and IdNo= IfNo and IdSr= IfSr   
  -- where IfPrtKey=@wCurrPrtn and not exists (select 'x' from Fgd where   
  --  FdPrtKey=IfPrtKey and FdIdNo=IfFdIdNo and FdCoCd= IfCoCd and FdTc= IfFdTc and FdYy= IfFdYy and   
  --  FdChr= IfFdChr and FdNo= IfFdNo and FdSr= IfFdSr and   
  --  FdBYy= IfFdBYy and FdBChr= IfFdBChr and FdBNo= IfFdBNo and   
  --  FdExpOdTc= IdExpOdTc and FdExpOdYy= IdExpOdYy and   
  --  FdExpOdChr= IdExpOdChr and FdExpOdNo= IdExpOdNo and   
  --  FdExpOdSr= IdExpOdSr)   
  --if @row>0   
  --begin  
  -- select @errmsg= 'Finished Goods Export Order Does Not Match With Invoice Design Export Voucher'  
  -- goto ErrHndlr  
  --end  
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
 End  
end  
  
 if @mode = 'D'   Begin  
  select @row=count('x') from deleted where 
   (exists(Select 'x' from InvRm where IrCoCd= IfCoCd and IrTc= IfTc and IrYy= IfYy and   
     IrChr= IfChr and IrNo= IfNo and IrSr= IfSr) Or  
   exists(Select 'x' from InvLab where iLCoCd= IfCoCd and iLTc= IfTc and iLYy= IfYy and   
     iLChr= IfChr and iLNo= IfNo and iLSr= IfSr))   
  if @row>0   
  begin  
   select @errmsg='Cannot Delete As Invoice Raw Material/ Labour Record Exists'  
   goto ErrHndlr  
  end  
  
 End  
  
 Return  
ErrHndlr:  
 Begin   select @errmsg='|'+@errmsg  
  raiserror(@errmsg,16,1)   rollback transaction   
 end  
end  
end


GOGO

ALTER TRIGGER [dbo].[InvHd_INUPDL] ON [dbo].[InvHd]    
    FOR INSERT,UPDATE,DELETE    NOT FOR REPLICATION
AS    
BEGIN    
 DECLARE @ROW INT    
 DECLARE @mode char(1)    
 DECLARE @errmsg varchar(255)    
 Declare @wHCd Varchar(8)     
 DECLARE @wTcTyp varchar(3)    
 DECLARE @wEntTyp varchar(8)    
 DECLARE @ReplAgt BIT      
 DECLARE @wCurrPrtn VarChar(1)     
Declare @wInFob Float  
 Declare @wInFobRs Float  
 Declare @wInLclOctroiRs Float  
 Declare @wInLclVatRs Float  
 Declare @wInLclCstRs Float    
  declare @wIsJLTc char(1)
     
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
     
    
 If  (@ReplAgt<>1)    
 Begin    
  if @mode in ('A','C')    
  Begin     
   If @mode='A' or (@mode='C' and (Update(InFOBRs) or Update(InFOB) or Update(InRsMulDiv) or Update(inCnvRt)))    
   Begin     
     select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT') 
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
 --5717
 /******** Only One TcTyp records can be Inserted at a time.********/     
 select @Row= Count(distinct(PValue)) from inserted Join param On    
  PTyp= 'TC' and PMCd= InTc and PScd= '' Where InPrtKey=@wCurrPrtn     
 If @Row> 1    
 begin    
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'    
    GoTo ErrHndlr    
 end    
 --5717   
 /******** Only One TcTyp records can be Deleted at a time. *********/     
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
  --5717
  If Update(InPrtKey) or Update(InIdNo) or Update(InCoCd) or Update(InTc) or Update(InYy) or Update(InChr) or Update(InNo)    
  Begin     
    Select @Row=Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and    
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
    select @Row=Count('x') from deleted del where del.InCls= 'Y' and     
      not exists (Select 'x' from inserted ins where 
       ins.InIdNo=del.InIdNo and ins.InCoCd= del.InCoCd     
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
    
    If @Row > 0    
    begin     
     Select @ErrMsg= 'Cannot Change InvExpNo,Cust/Contractor,Dt,LMG, LMP, LMS and LML Cst/Dt/Sal,CnvRt,FOB,Comm,Frt,Ins,Dsc,Deduct,Self Payable,TotVal,Dia/CS Cst,Adv Remit,Prn Dt, Inv Typ As Inv Is Clsd'    
     GoTo ErrHndlr    
    end    
    End    
        
  If Update(InDt) or Update(InVaDc) or Update(InVaX) or Update(InHRndOffFOB) or Update(InFrRmLoc) or Update(InToRmLoc)     
     or Update(InFrRmDc) or Update(InToRmDc) or Update(InHGldFiness) or Update(InHPlFiness)    
     or Update(InHSlFiness) or Update(InHPdFiness) or Update(InFobPer) or Update(InRemakingCtgs)    
      or Update(InInvGrsWtEqActual)    
  Begin     
    select @Row=Count('x') from deleted del where 
     ((@wTcTyp Not In ('FG', 'JT') and exists(Select 'x' from InvDsg where IdInIdNo=del.InIdNo and IdCoCd= del.InCoCd and IdTc= del.InTc     
       and IdYy= del.InYy and IdChr= del.InChr and IdNo= del.InNo)) Or    
      (@wTcTyp In ('FG', 'JT') and exists(Select 'x' from OrdMst where OmCoCd= del.InCoCd and OmTc= del.InTc     
       and OmYy= del.InYy and OmChr= del.InChr and OmNo= del.InNo))    
     ) and    
      not exists (Select 'x' from inserted ins where 
       ins.InIdNo=del.InIdNo and ins.InCoCd= del.InCoCd     
       and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
       and ins.InNo= del.InNo     
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
    If @Row > 0    
    begin     
     Select @ErrMsg= 'Cannot Chng Inv Dt, Gld, Plt, Slv And Pd Fineness, Value Addn For Dia, ColStn & Acc, FOB Rnd Opt, FOB %, Rm Locs, IR As ' + case When @wTcTyp In ('FG', 'JT') Then 'Ord' else 'Inv' End + 'Dsg Rec Exists'    
     GoTo ErrHndlr    
    end    
    End    
   --5717
    If @wTcTyp In ('CTB', 'FG')
  Begin    
    If Update(InLmgCst) or Update(InLmgDt) or Update(InLmgSal) Or    
       Update(InLmpCst) or Update(InLmpDt) or Update(InLmpSal) Or    
       Update(InLmsCst) or Update(InLmsDt) or Update(InLmsSal) Or    
       Update(InLmlCst) or Update(InLmlDt) or Update(InLmlSal)    
    Begin     
      select @Row=Count('x') from deleted del where del.InPrtKey=@wCurrPrtn and    
       ((@wTcTyp = 'CTB' and exists(Select 'x' from InvDsg where IdPRtKey=del.InPrtKey and IdInIdNo=del.InIdNo and IdCoCd= del.InCoCd and IdTc= del.InTc     
         and IdYy= del.InYy and IdChr= del.InChr and IdNo= del.InNo)) Or    
        (@wTcTyp In ('FG') and exists(Select 'x' from OrdMst where OmPRtKey=del.InPrtKey and OmCoCd= del.InCoCd and OmTc= del.InTc     
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
	     
   If @wTcTyp = 'IN'     
   Begin    
     If (Update(InLmgCst) or Update(InLmgDt) or Update(InLmgSal) Or    
         Update(InLmpCst) or Update(InLmpDt) or Update(InLmpSal) Or    
         Update(InLmsCst) or Update(InLmsDt) or Update(InLmsSal) Or    
         Update(InLmlCst) or Update(InLmlDt) or Update(InLmlSal) Or Update(InNotionRtYn))    
     Begin     
      select @Row=Count('x') from deleted del where @wTcTyp = 'IN'     
       and exists(Select 'x' from InvZRm where IzInIdNo=del.InIdNo and IzCoCd= del.InCoCd and IzTc= del.InTc     
          and IzYy= del.InYy and IzChr= del.InChr and IzNo= del.InNo) and     
        not exists (Select 'x' from inserted ins where 
         ins.InIdNo=del.InIdNo and ins.InCoCd= del.InCoCd     
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
  If @wTcTyp= 'IN'     
  Begin    
     If Update(InCmCd) or Update(InRsMulDiv) or Update(InCnvRt)    
     Begin     
      select @Row= Count('x') from deleted del where @wTcTyp= 'IN' and    
       (exists(Select 'x' from InvDsg where IdInIdNo=del.InIdNo and IdTc= del.InTc and IdYy= del.InYy and IdChr= del.InChr     
         and IdNo= del.InNo )     
       or exists(Select 'x' from InvFgd where IfInIdNo=del.InIdNo and IfTc= del.InTc and IfYy= del.InYy and IfChr= del.InChr     
         and IfNo= del.InNo )     
       )    
       And     
       not exists (Select 'x' from inserted ins where ins.InIdNo=del.InIdNo and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
         and ins.InNo= del.InNo and ins.InCmCd= del.InCmCd and ins.InMulDiv= del.InMulDiv     
         and convert (decimal (16, 4), ins.InCnvFct)= convert (decimal (16, 4), del.InCnvFct))     
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Cannot Change Customer Code, Currency Conversion Rate As Inv Dsg/ Bag Rec Exists'     
       GoTo ErrHndlr    
      end    
      End    
    End    
  If @wTcTyp In ('FG', 'JT')    
  Begin    
     If Update(InCmCd) or Update(InRsMulDiv) or Update(InCnvRt)    
     Begin     
      select @Row= Count('x') from deleted del where @wTcTyp In ('FG', 'JT') and     
       exists(Select 'x' from OrdMst where OmTc= del.InTc and OmYy= del.InYy and OmChr= del.InChr     
         and OmNo= del.InNo ) And     
       not exists (Select 'x' from inserted ins where ins.InIdNo=del.InIdNo and ins.InTc= del.InTc and ins.InYy= del.InYy and ins.InChr= del.InChr     
         and ins.InNo= del.InNo and ins.InCmCd= del.InCmCd and ins.InMulDiv= del.InMulDiv     
         and convert (decimal (16, 4), ins.InCnvFct)= convert (decimal (16, 4), del.InCnvFct))     
      If @Row> 0    
      begin     
       Select @ErrMsg= 'Cannot Change Customer Code, Currency Conversion Rate As jEmr Txn Exists'     
       GoTo ErrHndlr    
      end    
    End      
   End
   --5717
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
    
  If @mode='A' or (@mode='C' and (Update(InCoCd)))    
  Begin    
    select @row= count('x') from inserted where 
      not exists (select 'x' from Head where HCoCd= InCoCd and HCd= @wHCd)    
    if @row>0     
    begin     
     select @errMsg= 'Company Code Not Valid'    
     goto ErrHndlr    
    end    
    End    
        
  If @mode='A' or (@mode='C' and (Update(InTc)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from Param where PTyp= 'TC' and PMCd= InTc)     
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
        
  If @mode='A' or (@mode='C' and (Update(InExpNo)))    
  Begin    
    select @row= count('x') from inserted ins where ins.InPrtKey=@wCurrPrtn and    
     (Select count('x') from InvHd InHd where 
      InHd.InExpNo=ins.InExpNo)> 1     
     And Exists (Select 'x' From SysDb Where SysId='CRM')     
      if @row> 0     
      begin    
        select @errmsg= 'InExpNo Key Should Be Unique'    
        goto ErrHndlr    
      end    
    
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
        
  If @mode='A' or (@mode='C' and (Update(InCmCd)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and    
     not exists (select 'x' from CustMst where     
     CmCtg = InCmCtg and CmCd= InCmCd)
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
        
  If @mode='A' or (@mode='C' and (Update(InCnvRt)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and  convert( decimal (16, 4), InCnvRt)= convert( decimal (16, 4), 0)     
    if @row> 0     
    begin    
     select @errmsg= 'Rupee Conversion Rate Cannot Be Zero'    
     goto ErrHndlr    
    end    
    End    
        
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
  If @mode='A' or (@mode='C' and (Update(InCstCd)))    
  Begin    
    select @row= count('x') from inserted where InPrtKey=@wCurrPrtn and InCstCd<>'' and   
					not exists (select 'x' from CustStore where CstCmCtg =InCmCtg and CstCmCd =InCmCd and CstCd =InCstCd )
    if @row> 0     
    begin    
     select @errmsg= 'Customer Store Code Not Defined'    
     goto ErrHndlr    
    end
  End    
  
  If @mode='A' or (@mode='C' and (Update(InTot) or Update(InTotAs) or Update(InFOB) or Update(InHHideCommYN) or Update(InComm)    
                  or Update(InFrt) or Update(InIns) or Update(InHSelfDscYN) or Update(InDsc)    
                   or Update(InHSelfDeductYN) or Update(InDeduct) or Update(InTCS)))    
  Begin    
    Select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT')  and InPrtKey=@wCurrPrtn and     
     convert (decimal (16, 2), InTot)<>     
      (Case InLclInvYn when 'N' then     
        convert (decimal (16, 2), (case InTotAs when 'FOB' then InFOB +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End) +InTCS  
        when 'CIF' then InFOB+ (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeduct End) +InTCS
        when 'C&F' then InFOB+ (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InFrt +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End) +InTCS
        when 'C&I' then InFOB+ (case when InHHideCommYN ='Y' then 0 Else InComm End)+ InIns +(case when InHSelfDscYN ='Y' then 0 Else InDsc End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeduct End) +InTCS
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
  If @mode='A' or (@mode='C' and (Update(InTotRs) or Update(InTotAs) or Update(InFOBRs) or Update(InHHideCommYN) or Update(InCommRs)    
                  or Update(InFrtRs) or Update(InInsRs) or Update(InHSelfDscYN) or Update(InDscRs)    
                   or Update(InHSelfDeductYN) or Update(InDeductRs) or Update(InTCSRs)))    
  Begin    
    select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT') and InLclInvYn = 'N' and InPrtKEy=@wCurrPrtn and    
     (convert (decimal (16, 2), InTotRs))<>     
     convert (decimal (16, 2), (case InTotAs when 'FOB' then InFOBRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End) +InTCSRs
      when 'CIF' then InFOBRs+ (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) +(case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End) +InTCSRs
      when 'C&F' then InFOBRs+ (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InFrtRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End) +InTCSRs
      when 'C&I' then InFOBRs+ (case when InHHideCommYN ='Y' then 0 Else InCommRs End)+ InInsRs +(case when InHSelfDscYN ='Y' then 0 Else InDscRs End) + (case when InHSelfDeductYN ='Y' then 0 Else InDeductRs End) +InTCSRs
      else 0 end))        
    if @row> 0     
    begin    
     select @errmsg= 'Total Value In Rs Is Not Proper'    
     goto ErrHndlr    
    end    
    End    
        
  If @mode='A' or (@mode='C' and (Update(InLclInvYn) or Update(InTotRs) or Update(InFOB) or Update(InLclOctroiRs)     
                  or Update(InLclVatRs) or Update(InLclCstRs) or Update(InLclOthrValRs) or Update(InRsMulDiv)    
                   or Update(inCnvRt)))    
  Begin    
    select @row= count('x') from inserted where @wTcTyp in ('IN','FG', 'JT') and InLclInvYn = 'Y' and InPrtKey=@wCurrPrtn and    
     convert (decimal (16, 0), InTotRs)<> convert (decimal (16, 0),     
              (InFob+ InLclOctroiRs+ InLclVatRs+ InLclCstRs+ InLclOthrValRs)     
              * (case when InRsMulDiv= 'M' then inCnvRt else 1 end)     
            / (case when InRsMulDiv= 'D' then inCnvRt else 1 end) )     
    if @row> 0     
    begin    
     select @errmsg= 'Total Value In Rs Is Not Proper'    
     goto ErrHndlr    
    end     
    End    
        
   If @mode='A' or (@mode='C' and (Update(InFrRmDc) or Update(InToRmDc)))    
   Begin    
    Select @Row= Count('x') From Inserted Ins     
     Where (InFrRmDc <>'' and @wTcTyp in ('IN','FG','JT') and InPrtKey=@wCurrPrtn and    
      (Not Exists (Select 'x' From Param where PTyp= 'DC'    
       and PMCd = Ins.InFrRmDc and PSCd= '')))
	  or (InToRmDC<>'' and @wTcTyp in ('IN','FG','JT') and InPrtKey = @wCurrPrtn and 
	  (Not Exists (Select 'x' From Param where PTyp= 'DC'    
       and PMCd = Ins.InToRmDc and PSCd= '')))
	       
      If @Row> 0    
    begin    
     Select @ErrMsg= 'Invalid From/ To Rm DC'    
       GoTo ErrHndlr    
    end    
        
        
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
    
    Select @Row= Count('x') From Inserted Ins where @wTcTyp in ('IN','FG')  and InPrtKey=@wCurrPrtn     
     and Not Exists (Select 'x' From Loc where LocCoCd = InCoCd     
      and LocTyp= 'F' and (LocCd= InFrRmLoc or LocCd = InToRmLoc))    
        If @Row>0    
    begin    
     Select @ErrMsg= 'Atleast One RM Loc Must Be Of Type ''F'' '    
       GoTo ErrHndlr    
    end    
    End    
        
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
    
 if @mode='A'
 begin
   If @wTcTyp = 'FG' or @wTcTyp = 'JT'   
   Begin  
		  /*** Update InvHd ***/  
		  Update InvHd set  
			  @wInFob = convert (decimal (16, 2), InvHd.InFOB + IsNull((Select sum((ins.OdSalPrc*ins.OdOrdQty)+ins.OdJLVchVal) from OrdDsg ins where   
			   ins.OdPrtKey=InvHd.InPrtKey and ins.OdCoCd= InvHd.InCoCd and ins.OdTc= InvHd.InTc and ins.OdYy= InvHd.InYy   
			   and ins.OdChr= InvHd.InChr and ins.OdNo= InvHd.InNo), 0)),   
			  @wInFobRs= convert (decimal (16, 0), (@wInFob)  
				* (case when InvHd.InRsMulDiv= 'M' then InvHd.inCnvRt else 1 end)   
				/ (case when InvHd.InRsMulDiv= 'D' then InvHd.inCnvRt else 1 end) ),   
			  InFOB= @wInFob,   
			  InFOBRs= @wInFobRs,   
			  @wInLclOctroiRs= convert (decimal (16, 2), (Case when InvHd.InLclInvYn = 'Y' then   
				   (@wInFob * (InvHd.InCmOctroiPer / 100)) else 0 end)),  
			  @wInLclVatRs= convert (decimal (16, 2), (Case when InvHd.InLclInvYn = 'Y' then   
				   ((@wInFOB * (1 + (InvHd.InCmOctroiPer/100 ))+InvHd.InLclOthrValRs)* (InvHd.InHVatPer /100 )) else 0 end)),  
			  @wInLclCstRs= convert (decimal (16, 2), (Case when InvHd.InLclInvYn = 'Y' then   
				   ((@wInFOB * (1 + (InvHd.InCmOctroiPer/100 ))+InvHd.InLclOthrValRs)* (InvHd.InHCstPer /100 )) else 0 end)),  
			  InLclOctroiRs= @wInLclOctroiRs,   
			  InLclVatRs= @wInLclVatRs,   
			  InLclCstRs= @wInLclCstRs,   
			  InTot= convert (decimal (16, 2),   
				 (Case InvHd.InLclInvYn   
				 when 'N' then @wInFOB + (case InvHd.InTotAs when 'FOB' then (case when InvHd.InHSelfDscYN ='Y' then 0 Else InvHd.InDsc End) +(case when InvHd.InHSelfDeductYN ='Y' then 0 Else InvHd.InDeduct End)  
					   when 'CIF' then (case when InvHd.InHHideCommYN ='Y' then 0 Else InvHd.InComm End)+ InvHd.InFrt+ InvHd.InIns +(case when InvHd.InHSelfDscYN ='Y' then 0 Else InvHd.InDsc End) +(case when InvHd.InHSelfDeductYN ='Y' then 0 Else InvHd.InDeduct End)  
					   when 'C&F' then (case when InvHd.InHHideCommYN ='Y' then 0 Else InvHd.InComm End)+ InvHd.InFrt +(case when InvHd.InHSelfDscYN ='Y' then 0 Else InvHd.InDsc End) + (case when InvHd.InHSelfDeductYN ='Y' then 0 Else InvHd.InDeduct End)  
					   when 'C&I' then (case when InvHd.InHHideCommYN ='Y' then 0 Else InvHd.InComm End)+ InvHd.InIns +(case when InvHd.InHSelfDscYN ='Y' then 0 Else InvHd.InDsc End) + (case when InvHd.InHSelfDeductYN ='Y' then 0 Else InvHd.InDeduct End)  
					   else 0 end)  
				 when 'Y' then @wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InvHd.InLclOthrValRs   
				 else 0 end)),  
			  InTotRs= (Case InvHd.InLclInvYn   
				 when 'N' then @wInFOBRs + (case InvHd.InTotAs when 'FOB' then (case when InvHd.InHSelfDscYN ='Y' then 0 Else InvHd.InDscRs End) +(case when InvHd.InHSelfDeductYN ='Y' then 0 Else InvHd.InDeductRs End)  
						when 'CIF' then (case when InvHd.InHHideCommYN ='Y' then 0 Else InvHd.InCommRs End)+ InvHd.InFrtRs+ InvHd.InInsRs +(case when InvHd.InHSelfDscYN ='Y' then 0 Else InvHd.InDscRs End) +(case when InvHd.InHSelfDeductYN ='Y' then 0 Else InvHd.InDeductRs End)  
						when 'C&F' then (case when InvHd.InHHideCommYN ='Y' then 0 Else InvHd.InCommRs End)+ InvHd.InFrtRs +(case when InvHd.InHSelfDscYN ='Y' then 0 Else InvHd.InDscRs End) + (case when InvHd.InHSelfDeductYN ='Y' then 0 Else InvHd.InDeductRs End)  
						when 'C&I' then (case when InvHd.InHHideCommYN ='Y' then 0 Else InvHd.InCommRs End)+ InvHd.InInsRs +(case when InvHd.InHSelfDscYN ='Y' then 0 Else InvHd.InDscRs End) + (case when InvHd.InHSelfDeductYN ='Y' then 0 Else InvHd.InDeductRs End)  
						else 0 end)  
				 when 'Y' then convert (decimal (16, 0),   
				  (@wInFob+ @wInLclOctroiRs+ @wInLclVatRs+ @wInLclCstRs+ InvHd.InLclOthrValRs)   
				  * (case when InvHd.InRsMulDiv= 'M' then InvHd.inCnvRt else 1 end)   
				  / (case when InvHd.InRsMulDiv= 'D' then InvHd.inCnvRt else 1 end))  
				 else 0 end)  
		  from Inserted 
		  Join OrdDsg   on InPrtKey=OdPrtKey and InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
		  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo and OmPrtKey=OdPrtKey  
		  where 
		  Inserted.InPrtKey=InvHd.InPrtKey and Inserted.InCoCd =  InvHd.InCoCd and Inserted.InTc= InvHd.InTc and Inserted.InYy=InvHd.InYy   and Inserted.InChr=InvHd.InChr   and Inserted.InNo=  InvHd.InNo  
		  and Inserted.InPrtKey=OdPrtKey and Inserted.InCoCd= OdCoCd and Inserted.InTc= OdTc and Inserted.InYy= OdYy and Inserted.InChr= OdChr and Inserted.InNo= OdNo   
			  and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD ='FG' and PValue= OdTc)  
				or   
				  exists (Select 'x' from Param where PTyp= 'TCTYP' and   
					PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  
 end
 end
 return    
ErrHndlr:    
 begin    
  select @errmsg='|'+@errmsg   raiserror(@errmsg,16,1)    
  rollback transaction     
 end    
end    


GOGO


ALTER TRIGGER [dbo].[InvLab_INUPDL] ON [dbo].[InvLab] 
	    FOR INSERT,UPDATE,DELETE 
		NOT FOR REPLICATION
AS 
BEGIN 
	/* Trigger Version 3.02.0 - Speed Optimization */
	DECLARE @ROW INT 
	DECLARE @mode char(1) 
	DECLARE @errmsg varchar(255) 
	DECLARE @wTcTyp varchar(3)
	DECLARE @wEntTyp varchar(8)
	DECLARE @AddLabMCd varchar(5) 
	DECLARE @AddLabSCd varchar(8) 
	DECLARE @ReplAgt BIT  
	DECLARE @wCurrPrtn VarChar(1)
	DECLARE @wInCopyMode char(1) ='N'
	IF @@ROWCOUNT=0 RETURN 
 
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'InvLab' ) 
		RETURN 
 
	select @mode=case (select count('x') from inserted) when 0 then 'D' 
	else 
		case (select count('x') from deleted) when 0 then 'A' else 'C' end 
	end 
 
	Select @AddLabMCd= 'ZADD', @AddLabSCd= '-'

	
	If @Mode='D' 
	Begin 
		select @row= count('x') from deleted del where del.IlPrtKey<>@wCurrPrtn
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete Records From Previous Partition'
			goto ErrHndlr
		end
	End

	select @wInCopyMode = 'Y' where exists (select 'a' from inserted where Cpy='Y')

	if @wInCopyMode='N' 
	begin
	select @Row= Count(distinct(PValue)) from inserted Join param On 
		PTyp= 'TC' and PMCd= iLTc and PSCd= '' Where IlPrtKey=@wCurrPrtn
	If @Row> 1
	begin
		Select @ErrMsg= 'Transactions of only One TcType can be Acted On'
	  	GoTo ErrHndlr
	end

	select @Row= Count(distinct(PValue)) from deleted Join param on 
		PTyp= 'TC' and PMCd= iLTc and PSCd= '' Where IlPrtKey=@wCurrPrtn 
	If @Row> 1
	begin
		Select @ErrMsg= 'Transactions of only One TcType can be Acted on'
		GoTo ErrHndlr
	end
	end
	/*** Storing Entry Type In The Variable @wEntTyp ***/
	select @wEntTyp= PValue1 from Param  where PTyp= 'TCTYP' and PMCd= @wTcTyp and PSCd= ''
	--5717
	/******** Only One TcTyp records can be Inserted at a time. ********/	
	select @Row= Count(distinct(PValue)) from inserted Join param On 
		PTyp= 'TC' and PMCd= iLTc and PSCd= '' Where IlPrtKey=@wCurrPrtn
	If @Row> 1
	begin
		Select @ErrMsg= 'Transactions of only One TcType can be Acted On'
	  	GoTo ErrHndlr
	end
	--5717
	/******** Only One TcTyp records can be Deleted at a time. *********/	
	select @Row= Count(distinct(PValue)) from deleted Join param on 
		PTyp= 'TC' and PMCd= iLTc and PSCd= '' Where IlPrtKey=@wCurrPrtn 
	If @Row> 1
	begin
		Select @ErrMsg= 'Transactions of only One TcType can be Acted on'
		GoTo ErrHndlr
	end

 --	/******** Cannot Add/ Delete/ Change If Invoice Is Closed ********/ 
	if @wInCopyMode='N' 
 	begin
	if @mode in('A', 'C', 'D')
 	Begin
		/* Current Partition */
		select @Row= Count('x') from deleted join InvHd on InPrtKey=IlPrtKey and InCoCd= iLCoCd and InTc= iLTc and InYy= iLYy and InChr= iLChr and InNo= iLNo where IlPrtKey=@wCurrPrtn and InCls= 'Y' 
		If @Row > 0
		begin 
			Select @ErrMsg='Cannot Delete/ Change As Invoice Is Closed'
			GoTo ErrHndlr
		end

		/* Current Partition */
		select @Row=Count('x') from inserted join InvHd on InPrtKey=IlPrtKey and InCoCd= iLCoCd and InTc= iLTc and InYy= iLYy and InChr= iLChr and InNo= iLNo where IlPrtKey=@wCurrPrtn and InCls= 'Y' 
		If @Row > 0
		begin 
			Select @ErrMsg='Cannot Add/ Change As Invoice Is Closed'
			GoTo ErrHndlr
		end
	End
	end
	--5717
	if @mode = 'C'
 	Begin
		/******** Key Cannot Be Changed ********/ 
		/* Current Partition */
		select @Row=Count('x') from deleted del where del.IlPrtKey=@wCurrPrtn and
			 not exists (Select 'x' from inserted ins where ins.IlPrtKey=del.IlPrtKey and ins.iLCoCd= del.iLCoCd 
					and ins.iLTc= del.iLTc and ins.iLYy= del.iLYy and ins.iLChr= del.iLChr 
					and ins.iLNo= del.iLNo and ins.iLSr= del.iLSr and ins.iLSrNo= del.iLSrNo)
		If @Row > 0
		begin 
			Select @ErrMsg='Cannot Change Company Code/ Tc/Yy/Chr/No/Sr/SrNo'
			GoTo ErrHndlr
		end
	End

 --	if @mode = 'C'
 --	Begin
	--	/******** Key Cannot Be Changed ********/ 
	--	/* Current Partition */
	--	select @Row=Count('x') from deleted del where del.IlPrtKey=@wCurrPrtn and
	--		 not exists (Select 'x' from inserted ins where ins.IlPrtKey=del.IlPrtKey and ins.iLCoCd= del.iLCoCd 
	--				and ins.iLTc= del.iLTc and ins.iLYy= del.iLYy and ins.iLChr= del.iLChr 
	--				and ins.iLNo= del.iLNo and ins.iLSr= del.iLSr and ins.iLSrNo= del.iLSrNo)
	--	If @Row > 0
	--	begin 
	--		Select @ErrMsg='Cannot Change Company Code/ Tc/Yy/Chr/No/Sr/SrNo'
	--		GoTo ErrHndlr
	--	end
	--End

	if @wInCopyMode='N'
	begin
 	if @mode = 'A'
 	Begin
		/******** Only One Record Of The Lab Main Code 'ZADD' is Allowed For One InvSr Record ********/ 
		/* Current Partition */
		select @Row= Count('x') from inserted ins where IlPrtKey=@wCurrPrtn and iLMCd= @AddLabMCd and 
			 IsNull((Select count('x') from InvLab iL where Il.IlPrtKey=ins.IlPrtKey and iL.iLCoCd= ins.iLCoCd 
				and iL.iLTc= ins.iLTc and iL.iLYy= ins.iLYy and iL.iLChr= ins.iLChr 
				and iL.iLNo= ins.iLNo and iL.iLSr= ins.iLSr and iL.iLMCd= ins.iLMCd), 0)> 1 
		If @Row> 0
		begin 
			Select @ErrMsg= 'Only One Rec Of Labour MainCd ''+@AddLabMCd+'' is Allowed For One Inv Design Rec'
			GoTo ErrHndlr
		end

		/* Current Partition */
		select @row= count('x') from inserted where IlPrtKey=@wCurrPrtn and iLSCd= @AddLabMCd and iLSCd<> @AddLabSCd 
		if @row> 0 
		begin
			select @errmsg= 'Labour SubCd Should Always Be ''+@AddLabSCd+'' For The MainCd ''+@AddLabMCd+'''
			goto ErrHndlr
		end

		/*** (Jen 2.13) ***/
		/****** Cannot insert if record exists in InvZRm ******/
		/* Current Partition */
		select @row= count('x') from inserted where IlPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and 
			exists(Select 'x' from InvZRm where IzPrtKey=IlPrtKey and IzCoCd= IlCoCd and IzTc= IlTc and IzYy= IlYy 
				and IzChr= IlChr and IzNo= IlNo) 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Insert Record As Customs Matching Has Been Done For The Invoice'
			goto ErrHndlr
		end
		/*** (Jen 2.13) ***/


	End
	end
	if @wInCopyMode='N'
	begin
	if @mode in ('A','C')
	Begin

		/*** (ver 2.11) To Ensure that only Designs with Qty> 0 can have Rm recs ***/
		/* Current Partition */	
		select @row=count('x') from inserted where IlPrtKEy=@wCurrPrtn and not exists (select 'x' from InvDsg where 
				IdPrtKey=IlPrtKey and IdCoCd= iLCoCd and IdTc= iLTc and IdYy= iLYy and 
				IdChr= iLChr and IdNo= iLNo and IdSr= iLSr and convert(decimal(16, 1), IdQty)> 0)
		if @row>0 
		begin
			select @errmsg= 'Voucher Not Defined In Invoice Design/ Design Qty = 0'
			goto ErrHndlr
		end

		/* Current Partition */
		select @row=count('x') from inserted where IlPrtKey=@wCurrPrtn and iLSrNo<= 0 
		if @row>0 
		begin
			select @errmsg= 'Invoice Lab Serial Number Should Be > Zero'
			goto ErrHndlr
		end

		/* Current Partition */
		select @row=count('x') from inserted where IlPrtKEy=@wCurrPrtn and not exists (select 'x' from Param where PTyp= 'LABMCD' and PMCd= iLMCd)
		if @row>0 
		begin
			select @errmsg= 'Labour Main Code Not Defined'
			goto ErrHndlr
		end
		/* Current Partition */		select @row=count('x') from inserted where IlPrtKey=@wCurrPrtn and not exists (select 'x' from Param where PTyp= 'LABSCD' and PMCd= iLMCd and PSCd= iLSCd)
		if @row>0 
		begin
			select @errmsg= 'Labour Sub Code Not Defined'
			goto ErrHndlr
		end
		/****** iLQty >= 0 ******/
		/* Current Partition */
		select @row=count('x') from inserted where IlPrtKey=@wCurrPRtn and iLQty< 0 
		if @row>0 
		begin
			select @errmsg= 'Invoice Lab Qty Should Be >= Zero'
			goto ErrHndlr
		end

		/****** iLARt, iLAVal >= 0 ******/
		/* Current Partition */
		select @row=count('x') from inserted where IlPrtKey=@wCurrPrtn and (iLARt< 0 or iLAVal< 0) 
		if @row>0 
		begin
			select @errmsg= 'Actual Lab Rate/ Val Should Be >= Zero'
			goto ErrHndlr
		end

		/* Current Partition */
		select @row=count('x') from inserted where IlPrtKey=@wCurrPrtn and not exists (select 'x' from Param where PTyp= 'LABQW' and PMCd= iLQw)
		if @row>0 
		begin
			select @errmsg= 'Lab Q/W Not Defined  (Mention If Labour Value Calculation Is By Quantity Or Weight)'
			goto ErrHndlr
		end
	End
	end

	/*** (Jen 2.13) ***/
	if @mode = 'D' 	Begin
		/****** Cannot delete if record exists in InvZRm ******/
		/* Current Partition */
		select @row=count('x') from deleted where IlPrtKEy=@wCurrPrtn and
			exists(Select 'x' from InvZRm where IzPrtKey=IlPrtKey and IzCoCd= IlCoCd and IzTc= IlTc 
				and IzYy= IlYy and IzChr= IlChr and IzNo= IlNo) 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Customs Matching Has Been Done For The Invoice'
			goto ErrHndlr
		end
	End
	/*** (Jen 2.13) ***/


	return

ErrHndlr:
	begin		select @errmsg='|'+@errmsg
		raiserror(@errmsg,16,1)		rollback transaction 
	end
end


GOGO


ALTER TRIGGER [dbo].[InvRm_INUPDL] ON [dbo].[InvRm]       
   FOR INSERT,UPDATE,DELETE       
	NOT FOR REPLICATION
AS       
BEGIN       
 /* Trigger Version 3.11.1 */      
 DECLARE @ROW INT       
 DECLARE @mode char(1)       
 DECLARE @errmsg varchar(255)       
 DECLARE @wTcTyp varchar(3)      
 DECLARE @wEntTyp varchar(8)      
 DECLARE @ReplAgt BIT        
 DECLARE @wCurrPrtn VarChar(1)      
 DECLARE @wInCopyMode char(1) ='N'
       
 IF @@ROWCOUNT=0 RETURN       
       
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'InvRm' )       
  RETURN       
       
 select @mode=case (select count('x') from inserted) when 0 then 'D'       
 else       
  case (select count('x') from deleted) when 0 then 'A' else 'C' end       
 end       
 select @wInCopyMode = 'Y' where exists (select 'a' from inserted where Cpy='Y')

     
 /* Sachin 3.02 - Constant For Current Partition */      
 Select @wCurrPrtn = 'C'       
      
 /* **** Checking Whether a Normal Transaction or Replicated One **** */      
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))      
      
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */       
       
 If  (@ReplAgt<>1)      
 Begin      
      
  --5717
 /******** Only One TcTyp records can be Inserted at a time.********/       
 select @Row= Count(distinct(PValue)) from inserted Join param On      
  PTyp= 'TC' and PMCd= IrTc and PScd= '' Where IrPrtKey=@wCurrPrtn       
 If @Row> 1      
 begin      
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'      
    GoTo ErrHndlr      
 end       
 /******** Only One TcTyp records can be Deleted at a time. *********/       
 select @Row= Count(distinct(PValue)) from deleted Join param on       
  PTyp= 'TC' and PMcd= IrTc and PSCd= '' Where IrPrtKey=@wCurrPrtn      
 If @Row> 1      
 begin      
  Select @ErrMsg= 'Transactions of only One TcType can be Acted on'      
  GoTo ErrHndlr      
 end      
     
      
 /*** Storing Tc Type In The Variable @wTcTyp ***/       
 if @mode in ('A','C')      
 begin      
  select @wTcTyp= max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= IrTc and PSCd= ''      
 end      
      
 if @mode in ('D')      
 begin      
  select @wTcTyp= max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= IrTc and PSCd= ''      
 end      
      
 /*** Storing Entry Type In The Variable @wEntTyp ***/      
 select @wEntTyp= PValue1 from Param  where PTyp= 'TCTYP' and PMCd= @wTcTyp and PSCd= ''      

      
  /******** Cannot Add/ Delete/ Change If Invoice Is Closed ********/       
if @wInCopyMode='N' 
  begin
  if @mode in('A', 'C', 'D')      
  Begin      
  select @Row=Count('x') from deleted       
   join InvHd on InPrtKey=IrPrtKey and InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo       
  where IrPrtKEy=@wCurrPrtn and InCls= 'Y'       
  If @Row > 0      
  begin       
   Select @ErrMsg='Cannot Delete/ Change As Invoice Is Closed'      
   GoTo ErrHndlr      
  end      
      
  /* Current Partition */      
  select @Row=Count('x') from inserted       
   join InvHd on InPrtKey=IrPrtKey and InCoCd= IrCoCd and InTc= IrTc and InYy= IrYy and InChr= IrChr and InNo= IrNo       
  where IrPrtKey=@wCurrPrtn and InCls= 'Y'       
  If @Row > 0      
  begin       
   Select @ErrMsg='Cannot Add/ Change As Invoice Is Closed'      
   GoTo ErrHndlr      
  end      
 End      
end      


 if @mode = 'A'   
 Begin      
  /****** Cannot insert if record exists in InvZRm with IrRmZCd = IzRmZCd ******/      
  /* Current Partition */      
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and       
   exists(Select 'x' from InvZRm where IzPrtKey=IrPrtKey and IzRmZCd= IrRmZCd and IzCoCd= IrCoCd       
    and IzTc= IrTc and IzYy= IrYy and IzChr= IrChr and IzNo= IrNo)       
  if @row>0       
  begin      
   select @errmsg='Cannot Insert Record As Customs Matching Has Been Done For The Rm Code'      
   goto ErrHndlr      
  end      
 End      
      
 if @wInCopyMode='N' 
 begin    
  if @mode = 'C'      
  Begin      
  --5717
  select @Row=Count('x') from deleted del where del.IrPrtKEy=@wCurrPrtn and      
    not exists (Select 'x' from inserted ins where ins.IrPrtKey=del.IrPRtKey and ins.IrIdNo=del.IrIdNo       
     and ins.IrCoCd= del.IrCoCd and ins.IrTc= del.IrTc and ins.IrYy= del.IrYy and ins.IrChr= del.IrChr       
     and ins.IrNo= del.IrNo and ins.IrSr= del.IrSr and ins.IrSrNo= del.IrSrNo)      
  If @Row > 0      
  begin       
   Select @ErrMsg='Cannot Change Company Code/ Tc/Yy/Chr/No/Sr/SrNo'      
   GoTo ErrHndlr      
  end         
      
  select @Row=Count('x') from deleted del where not exists (Select 'x' from inserted ins       
    where ins.IrIdNo=del.IrIdNo and ins.IrCoCd= del.IrCoCd       
     and ins.IrTc= del.IrTc and ins.IrYy= del.IrYy and ins.IrChr= del.IrChr       
     and ins.IrNo= del.IrNo and ins.IrSr= del.IrSr and ins.IrSrNo= del.IrSrNo       
     and ins.IrRmCd= del.IrRmCd and ins.IrRmZCd= del.IrRmZCd       
     and ins.IrFrRmLoc= del.IrFrRmLoc and ins.IrFrRmDc= del.IrFrRmDc       
     and ins.IrToRmLoc= del.IrToRmLoc and ins.IrToRmDc= del.IrToRmDc       
     and ins.IrRMSz=Del.IrRmSz      and ins.IrRMSz2=Del.IrRmSz2 and      ins.IrRMSz3=Del.IrRmSz3 and ins.IrRMSz2=Del.IrRmSz2 and ins.IrRMSz3=Del.IrRmSz3 and ins.IrStkRt=del.IrStkRt and ins.IrLotNo=del.IrLotNo)      
  If @Row > 0      
  begin       
   Select @ErrMsg= 'Cannot Change Rm Code/ Customs Code/ Rm Locations/ Rm DC/ Rm Sz/ Stk Rt/ Lot No'      
   GoTo ErrHndlr      
  end      
  /******** If @wTcTyp= 'IN' then Cannot Change IrRmZWt, IrRmQty, IrRmZRt if record exists in InvZRm for RmZCd= IzRmZCd ********/       
  /* Current Partition */      
  select @Row=Count('x') from deleted del where IrPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and       
    not exists (Select 'x' from inserted ins where ins.IrPRtkey=del.IrPrtKey       
     and ins.IrIdNo=del.IrIdNo and ins.IrCoCd= del.IrCoCd       
     and ins.IrTc= del.IrTc and ins.IrYy= del.IrYy and ins.IrChr= del.IrChr       
     and ins.IrNo= del.IrNo and ins.IrSr= del.IrSr and ins.IrSrNo= del.IrSrNo       
     and ins.IrRmZWt= del.IrRmZWt and ins.IrRmQty= del.IrRmQty       
     and ins.IrRmZRt= del.IrRmZRt)       
    and exists (Select 'x' from InvZRm where IzPrtKey=del.IrPrtKey and IzCoCd= del.IrCoCd and IzTc= del.IrTc       
     and IzYy= del.IrYy and IzChr= del.IrChr and IzNo= del.IrNo       
     and IzRmZCd= del.IrRmZCd)      
  If @Row > 0      
  begin       
   Select @ErrMsg='Cannot Change Customs Wt/ Qty/ Rt As Matching Has Already Been Done For This Custom Code'      
   GoTo ErrHndlr      
  end      
 End      
 end     
if @wInCopyMode='N' 
begin      
 if @mode in ('A','C')      
 begin      
  select @row= count('x') from inserted where IrPrtKey=@wCurrPrtn and not exists (select 'x' from InvDsg where       
    IdPRtKey=IrPrtKey and IdIdNo=IrIdIdNo and IdCoCd= IrCoCd and IdTc= IrTc and IdYy= IrYy and IdChr= IrChr and IdNo= IrNo       
    and IdSr= IrSr and convert(decimal(16, 1), IdQty)> 0)       
  if @row> 0       
  begin      
   select @errmsg= 'Voucher Not Defined In Invoice Design/ Design Qty = 0'      
   goto ErrHndlr      
  end      
  --5717
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and IrSrNo<= 0       
  if @row>0       
  begin      
   select @errmsg= 'Invoice Rm Serial Number Should Be > Zero'      
   goto ErrHndlr      
  end 
  --5717
  Select @Row=Count('x') From Inserted Ins where IrPrtKey=@wCurrPrtn and IrLotNo<> '' and       
   Not Exists (Select 'x' from SYyMm Join Loc On LocTyp='F' and LocCd=SLoc where SPrtKey=IrPRtKey and SCoCd= IrCoCd /*and SLoc= 'FG'*/ and SLotNo= IrLotNo)       
    If @Row>0      
  begin      
   Select @ErrMsg= 'Invalid Lot No.'      
   GoTo ErrHndlr      
  end  
  --5717
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and IrRmZCd<> '' and not exists (select 'x' from RmMst where RmZ= 'Y' and RmCd= IrRmZCd)      
  if @row>0       
  begin      
   select @errmsg= 'Invoice Customs Rm Code Not Defined'      
   goto ErrHndlr      
  end 
  --5717
  Select @Row= Count('x') From Inserted Ins Where Ins.IrPrtKey=@wCurrPrtn and  @wTcTyp= 'IN' and       
   ( Not Exists (Select 'x' From Param where PTyp= 'DC'      
     and PMCd = Ins.IrFrRmDc and PSCd= '')      
    or Not Exists (Select 'x' From Param where PTyp= 'DC'      
     and PMCd = Ins.IrToRmDc and PSCd= '')      
   )      
    If @Row> 0      
  begin      
   Select @ErrMsg= 'Invalid From/ To Rm DC'      
     GoTo ErrHndlr      
  end   
  --5717
  Select @Row= Count('x') From Inserted Ins Where Ins.IrPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and       
   ( (IrToRmDc= IrFrRmDc And IrFrRmDc<> '')        
    Or (IrFrRmDc<> '' and IrToRmDc= '')      
   )      
    If @Row> 0      
  begin      
   Select @ErrMsg= 'Invalid To Rm Dc'      
   GoTo ErrHndlr      
  end   
      
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and not exists (select 'x' from RmMst where RmZ= 'N' and RmPrtKey=IrPrtKey and RmCd= IrRmCd)      
  if @row>0       
  begin      
   select @errmsg= 'Invoice Rm Code Not Defined'      
   goto ErrHndlr      
  end      
      
  Select @Row=Count('x') From Inserted Ins      
   Join InvDsg On IdIdNo=IrIdIdNo 
  Where Not Exists (Select 'x' from OrdRm where OrCoCd= IdCoCd       
   and OrTc= IdExpOdTc and OrYy= IdExpOdYy and OrChr= IdExpOdChr      
   and OrNo= IdExpOdNo and OrSr= IdExpOdSr and OrRmCd= IrRmCd)      
    If @Row>0      
  begin      
   Select @ErrMsg= 'Rm Code Does Not Exist In Order'      
   GoTo ErrHndlr      
  end      
      
  --/* Zubin 212 (Lot No shld be valid from SYyMm for Loc= 'FG') */       
  /* Current Partition */      
  --/* **** Manali 20/03/09 - Check changed - Lot No shld be valid from SYyMm for LocTyp= 'F' (Diajewel) */      
  --Select @Row=Count('x') From Inserted Ins where IrPrtKey=@wCurrPrtn and IrLotNo<> '' and       
  -- Not Exists (Select 'x' from SYyMm Join Loc On LocTyp='F' and LocCd=SLoc where SPrtKey=IrPRtKey and SCoCd= IrCoCd /*and SLoc= 'FG'*/ and SLotNo= IrLotNo)       
  --  If @Row>0      
  --begin      
  -- Select @ErrMsg= 'Invalid Lot No.'      
  -- GoTo ErrHndlr      
  --end        
      
      
  --/****** RmQty >= 0 ******/      
  --/* Current Partition */      
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and IrRmQty< 0       
  if @row>0       
  begin      
   select @errmsg= 'Invoice Rm Qty Should Be >= Zero'      
   goto ErrHndlr      
  end      
      
  --/****** Actual Wt, Rt, Val, Set Rt, Set Val >= 0 ******/   
  --/* Current Partition */      
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and (IrRmAWt< 0 or IrRmARt< 0 or IrRmAVal< 0 or IrSetARt< 0 or IrSetAVal< 0)       
  if @row>0       
  begin      
   select @errmsg= 'Actual Weight/ Rate/ Value Should Be >= Zero'      
   goto ErrHndlr      
  end      
      
  --/****** Invoice Wt, Rt, Val, Set Rt, Set Val  >= 0 ******/      
  --/* Current Partition */      
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and (IrRmIWt< 0 or IrRmIRt< 0 or IrRmIVal< 0 or IrSetIRt< 0 or IrSetIVal< 0)       
  if @row>0       
  begin      
   select @errmsg= 'Invoice Weight/ Rate/ Value Should Be >= Zero'      
   goto ErrHndlr      
  end      
  -- /****** Customs Wt, Rt, Val, Set Rt, Set Val  >= 0 ******/      
  --/* Current Partition */      
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and (IrRmZWt< 0 or IrRmZRt< 0 or IrRmZVal< 0 or IrSetZRt< 0 or IrSetZVal< 0)      
  if @row>0       
  begin      
   select @errmsg= 'Customs Weight/ Rate/ Value Should Be >= Zero'      
   goto ErrHndlr      
  end      
      
  /* Current Partition */      
  
  select @row= count('x') from inserted join RmMst on RmCd= IrRmCd and RmPrtKey=IrPRtKey      
   where IrPrtKey=@wCurrPrtn and (RmCtg in ('D', 'C') or IrSetSCd<> '') and not exists (select 'x' from Param where PTyp= 'LABSCD' and PMCd= 'SET'       
    and PSCd= IrSetSCd)      
  if @row> 0       
  begin      
   select @errmsg= 'Setting Code Not Defined'      
   goto ErrHndlr      
  end      
  /* Current Partition */      
  /* **** Manali 3.03 - 21/06/08 - Pd Mod - 'L' added ****** */      
  select @row=count('x') from inserted join RmMst on RmCd= IrRmCd and RmPrtKey=IrPrtKey      
   where IrPrtKey=@wCurrPrtn and RmCtg in ('G', 'P', 'M', 'S', 'L') and (IrSetSCd<> '' or IrSetARt<> 0 or IrSetiRt<> 0 or IrSetZRt<> 0 or       
    IrSetAVal<> 0 or IrSetiVal<> 0 or IrSetZVal<> 0)         
  if @row> 0       
  begin      
   select @errmsg='Setting Code/ Rate/ Value Should Not Be Entered For This Rm Code'      
   goto ErrHndlr      
  end
  
   /* Current Partition */   select @row= count('x') from inserted where IrPrtKey=@wCurrPrtn and not exists (select 'x' from Param where PTyp= 'YN' and PMCd= IrMainMet)       
  if @row> 0       
  begin      
   select @errmsg= 'Enter Main Metal (Y/ N)'      
   goto ErrHndlr      
  end      
  /****** Only One Main Metal Record Is allowed Per InvSr ******/      
  /* Current Partition */      
  select @row=count('x') from inserted ins where ins.IrPrtKey=@wCurrPrtn and      
   (select count('x') from InvRm IRM where ins.IrPrtKey=IRM.IrPrtKey       
   and ins.IrIdIdNo=IRM.IrIdIdNo and ins.IrCoCd= IRM.IrCoCd and       
   ins.IrTc= IRM.IrTc and ins.IrYy= IRM.IrYy and ins.IrChr= IRM.IrChr and ins.IrNo= IRM.IrNo and       
   ins.IrSr= IRM.IrSr and IRM.IrMainMet = 'Y') > 1       
  if @row> 0       
  begin      
   select @errmsg='Cannot Have More Than One Main Metal'      
   goto ErrHndlr      
  end
      
  /* Current Partition */       
/*  select @row=count('x') from inserted where IrRmZCd<> '' and not exists (select 'x' from RmMst where RmZ= 'Y' and RmCd= IrRmZCd)      
  if @row>0       
  begin      
   select @errmsg= 'Invoice Customs Rm Code Not Defined'      
   goto ErrHndlr      
  end      
  */    
  /* Current Partition */      
  select @row=count('x') from inserted where IrPrtKey=@wCurrPrtn and IrRmZCd<> '' and       
   not exists (select 'x' from RmMst where RmZ= 'N' and RmPrtKey=IrPrtKey and RmCd= IrRmCd and RmZCd= IrRmZCd)      
  if @row>0       
  begin      
   select @errmsg= 'Invoice Customs Rm Code Not Defined For The Rm Code'      
   goto ErrHndlr      
  end      
  /****** IrFrRmLoc or IrToRmLoc Should Have LocTyp= 'XS' ********/      
  /* Current Partition */      
  Select @Row= Count('x') From Inserted Ins Where IrPRtKey=@wCurrPrtn and @wTcTyp= 'IN' and       
   Not Exists (Select 'x' From Loc where LocCoCd = IrCoCd       
    and LocTyp= 'XS' and (LocCd= IrFrRmLoc or LocCd = IrToRmLoc))      
    If @Row> 0      
  begin      
   Select @ErrMsg= 'Atleast One RM Loc Must Be Of Type ''XS'' '      
     GoTo ErrHndlr      
  end      
  /****** IrFrRmLoc or IrToRmLoc Should Have LocTyp= 'F' ********/      
  /* Current Partition */      
Select @Row= Count('x') From Inserted Ins Where IrPrtKEy=@wCurrPrtn and @wTcTyp= 'IN' and       
   Not Exists (Select 'x' From Loc where LocCoCd = IrCoCd       
    and LocTyp= 'F' and (LocCd= IrFrRmLoc or LocCd = IrToRmLoc))      
      If @Row>0      
  begin      
   Select @ErrMsg= 'Atleast One RM Loc Must Be Of Type ''F'' '      
     GoTo ErrHndlr      
  end      
  /****** RM Wt/ Qty Can Be Issued From Locations Of Type 'F' Only ********/      
  /* Current Partition */      
  Select @Row= Count('x') From Inserted Ins where Ins.IrPrtkey=@wCurrPrtn and @wTcTyp= 'IN' and       
   Not Exists (Select 'x' From Loc where LocCoCd = IrCoCd and       
    LocTyp= 'F' and ( (LocCd= IrFrRmLoc and IrFrRmDc= 'C')  or       
       (LocCd= IrToRmLoc and IrToRmDc= 'C')               ))      
      If @Row> 0      
  begin      
   Select @ErrMsg= 'RM Wt/ Qty Can Be Issued From Locations Of Type ''F'' Only'      
     GoTo ErrHndlr      
  end      
 end      
  end    
  if @mode = 'D'   
  Begin      
  select @row=count('x') from deleted where IrPrtKey=@wCurrPrtn and      
   exists(Select 'x' from InvZRm where IzPrtKey=IrPrtKey and IzCoCd= IrCoCd and IzTc= IrTc       
    and IzYy= IrYy and IzChr= IrChr and IzNo= IrNo and IzRmZCd= IrRmZCd)       
  if @row>0       
  begin      
   select @errmsg='Cannot Delete As Customs Matching Has Been Done For The Rm Code'      
   goto ErrHndlr      
  end      
 End      
      
 if @mode in ('C','D')      
 begin      
  /* Current Partition */      
  update SYyMm set       
  SRmCrQty= SRmCrQty- (select isnull(sum(Del.IrRmQty), 0) from Deleted Del       
   where Del.IrFrRmDc= 'C' and Del.IrPrtKey=SPrtKey and Del.IrCoCd= SCoCd and Del.IrRmCd= SRmCd       
   and Del.IrLotNo= SLotNo and SRmSz= Del.IrRmSz and SRmSz2= Del.IrRmSz2 and SRmSz3= Del.IrRmSz3 and SRmStkRt= Del.IrStkRt and Del.IrFrRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))= SYyMm),      
  SRmCrWt= SRmCrWt-(select isnull(sum(Del.IrRmAWt),0) from Deleted Del       
    where Del.IrFrRmDc='C' and Del.IrPrtKey=SPrtKey and Del.IrCoCd= SCoCd and Del.IrRmCd= SRmCd       
   and Del.IrLotNo= SLotNo and SRmSz= Del.IrRmSz and SRmSz2= Del.IrRmSz2 and SRmSz3= Del.IrRmSz3 and SRmStkRt= Del.IrStkRt and Del.IrFrRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmDrQty= SRmDrQty-(select isnull(sum(Del.IrRmQty),0) from Deleted Del       
   where Del.IrFrRmDc='D' and Del.IrPRtKey=SPrtKey and Del.IrCoCd= SCoCd and Del.IrRmCd= SRmCd       
   and Del.IrLotNo= SLotNo and SRmSz= Del.IrRmSz and SRmSz2= Del.IrRmSz2 and SRmSz3= Del.IrRmSz3 and SRmStkRt= Del.IrStkRt and Del.IrFrRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmDrWt= SRmDrWt-(select isnull(sum(Del.IrRmAWt),0) from Deleted Del       
   where Del.IrFrRmDc='D' and Del.IrPrtKey=SPrtKey and Del.IrCoCd= SCoCd and Del.IrRmCd= SRmCd       
   and Del.IrLotNo= SLotNo and SRmSz= Del.IrRmSz and SRmSz2= Del.IrRmSz2 and SRmSz3= Del.IrRmSz3 and SRmStkRt= Del.IrStkRt and Del.IrFrRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm)      
  from Deleted        
   where @wTcTyp= 'IN' and Deleted.IrPrtKey=@wCurrPrtn and Deleted.IrPrtKey=SPrtKey and Deleted.IrCoCd= SCoCd       
   and Deleted.IrRmCd= SRmCd and Deleted.IrLotNo= SLotNo and SRmSz= Deleted.IrRmSz  and SRmSz2= Deleted.IrRmSz2            
   and SRmStkRt= Deleted.IrStkRt  and Deleted.IrFrRmLoc= SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm      
   and (Deleted.IrFrRmDc='D'  or Deleted.IrFrRmDc='C') and       
   ((@Mode='C' and Exists(Select 'x' From Inserted ins where ins.IrCoCd =Deleted.IrCoCd and ins.IrTc=Deleted.IrTc and ins.IrYy=Deleted.IrYy and      
         ins.IrChr=Deleted.IrChr and ins.IrNo=Deleted.IrNo and ins.IrSr=Deleted.IrSr and ins.IrSrNo=Deleted.IrSrNo and          
      (ins.IrPrtKey<>Deleted.IrPrtKey Or  ins.IrCoCd<>Deleted.IrCoCd Or ins.IrRmCd<>Deleted.IrRmCd Or ins.IrLotNo<>Deleted.IrLotNo      
        Or  ins.IrRmSz<>Deleted.IrRmSz or ins.IrRmSz2<>Deleted.IrRmSz2 or   ins.IrRmSz3<>Deleted.IrRmSz3 or  ins.IrStkRt<>Deleted.IrStkRt Or ins.IrFrRmLoc<>Deleted.IrFrRmLoc Or       
        Ins.IrRmAWt<>Deleted.IrRmAWt Or Ins.IrRmQty<>Deleted.IrRmQty Or ins.IrFrRmDc<>Deleted.IrFrRmDc Or      
         convert(int,convert(varchar(6),ins.IrDt,112))<>convert(int,convert(varchar(6),Deleted.IrDt,112))))) Or @Mode='D')      
      
             
 end      
      
      
 if @mode in ('A','C')      
 begin      
  /* Current Partition */      
  insert into SYyMm(SPrtKey,SCoCd, SRmCd, SLotNo, SRmSz, SRmSz2, SRmSz3, SRmStkRt, SLoc, SYyMm, SRmDrQty, SRmDrWt, SRmCrQty, SRmCrWt)       
   select distinct IrPrtKey, IrCoCd, IrRmCd, IrLotNo, IrRmSz, IrRmSz2,IrRmSz3,IrStkRt, IrFrRmLoc, convert(int,convert(varchar(6),IrDt,112)), 0, 0, 0, 0      
   from inserted      
   where @wTcTyp= 'IN' and not exists(select 'x' from SYyMm where       
    SPrtKey=IrPrtKey and SCoCd= IrCoCd and SRmCd= IrRmCd and SLotNo= IrLotNo and SRmSz= IrRmSz 
	and SRmSz2= IrRmSz2 and SRmSz3= IrRmSz3                    
    and SRmStkRt= IrStkRt and SLoc=IrFrRmLoc       
    and SYyMm= convert(int,convert(varchar(6),IrDt,112)))       
      
  /* Current Partition */      
  update SYyMm set      
  SRmCrQty= SRmCrQty+ (select isnull(sum(Ins.IrRmQty),0) from Inserted Ins       
   where Ins.IrFrRmDc='C' and Ins.IrPrtKey=SPrtKey and Ins.IrCoCd= SCoCd and Ins.IrRmCd=SRmCd and Ins.IrLotNo=SLotNo       
   and SRmSz= Ins.IrRmSz and SRmSz2= Ins.IrRmSz2 and SRmSz3= Ins.IrRmSz3 and SRmStkRt= Ins.IrStkRt and Ins.IrFrRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmCrWt= SRmCrWt+ (select isnull(sum(Ins.IrRmAWt),0) from Inserted Ins      
   where Ins.IrFrRmDc='C' and Ins.IrPrtKey=SPrtKey and Ins.IrCoCd= SCoCd and Ins.IrRmCd=SRmCd and Ins.IrLotNo=SLotNo       
   and SRmSz= Ins.IrRmSz and SRmSz2= Ins.IrRmSz2 and SRmSz3= Ins.IrRmSz3 and SRmStkRt= Ins.IrStkRt and Ins.IrFrRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmDrQty= SRmDrQty+ (select isnull(sum(Ins.IrRmQty),0) from Inserted Ins      
   where Ins.IrFrRmDc='D' and Ins.IrPrtKey=SPrtKey and Ins.IrCoCd= SCoCd and Ins.IrRmCd=SRmCd and Ins.IrLotNo=SLotNo       
   and SRmSz= Ins.IrRmSz    and SRmSz2= Ins.IrRmSz2 and SRmSz3= Ins.IrRmSz3 and SRmStkRt= Ins.IrStkRt and Ins.IrFrRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmDrWt= SRmDrWt+ (select isnull(sum(Ins.IrRmAWt),0) from Inserted Ins      
   where Ins.IrFrRmDc='D'  and Ins.IrCoCd= SCoCd and Ins.IrRmCd=SRmCd and Ins.IrLotNo=SLotNo       
   and SRmSz= Ins.IrRmSz    and SRmSz2= Ins.IrRmSz2    and SRmSz3= Ins.IrRmSz3 and SRmStkRt= Ins.IrStkRt and Ins.IrFrRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm)      
  from Inserted      
  where Inserted.IrPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and Inserted.IrPRtKey=SPrtKey and Inserted.IrCoCd= SCoCd       
   and Inserted.IrRmCd= SRmCd and Inserted.IrLotNo= SLotNo and SRmSz= Inserted.IrRmSz       and SRmSz2= Inserted.IrRmSz2 and SRmSz3= Inserted.IrRmSz3              
   and SRmStkRt= Inserted.IrStkRt and Inserted.IrFrRmLoc= SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm       
   and (Inserted.IrFrRmDc='D' or Inserted.IrFrRmDc='C') and       
   ((@Mode='C' and Exists(Select 'x' From Deleted Del where inserted.IrCoCd =Del.IrCoCd and inserted.IrTc=Del.IrTc and inserted.IrYy=Del.IrYy and      
        inserted.IrChr=Del.IrChr and inserted.IrNo=Del.IrNo and inserted.IrSr=Del.IrSr and inserted.IrSrNo=Del.IrSrNo and          
     (inserted.IrPrtKey<>Del.IrPrtKey Or  inserted.IrCoCd<>Del.IrCoCd Or inserted.IrRmCd<>Del.IrRmCd Or inserted.IrLotNo<>Del.IrLotNo      
       Or  inserted.IrRmSz<>Del.IrRmSz Or  inserted.IrRmSz2<>Del.IrRmSz2 Or  inserted.IrRmSz3<>Del.IrRmSz3 or  inserted.IrStkRt<>Del.IrStkRt Or inserted.IrFrRmLoc<>Del.IrFrRmLoc Or       
       inserted.IrRmAWt<>Del.IrRmAWt Or inserted.IrRmQty<>Del.IrRmQty Or inserted.IrFrRmDc<>Del.IrFrRmDc Or       
        convert(int,convert(varchar(6),inserted.IrDt,112))<>convert(int,convert(varchar(6),Del.IrDt,112))))) Or @Mode='A')      
      
 end       
       
      
 /*** Update SYyMm from TdToRmLoc ***/      
 if @mode in ('C','D')      
 begin      
  /* Current Partition */      
  update SYyMm set      
  SRmCrQty=SRmCrQty-(select isnull(sum(Del.IrRmQty),0) from Deleted Del      
   where Del.IrToRmDc='C' and Del.IrPrtKey=SPrtKey and Del.IrCoCd= SCoCd and Del.IrRmCd=SRmCd and Del.IrLotNo=SLotNo       
   and SRmSz= Del.IrRmSz    and SRmSz2= Del.IrRmSz2    and SRmSz3= Del.IrRmSz3 and SRmStkRt= Del.IrStkRt and Del.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmCrWt=SRmCrWt-(select isnull(sum(Del.IrRmAWt),0) from Deleted Del       
   where Del.IrToRmDc='C' and Del.IrPrtKey=SPrtKey and Del.IrCoCd= SCoCd and Del.IrRmCd=SRmCd and Del.IrLotNo=SLotNo       
   and SRmSz= Del.IrRmSz    and SRmSz2= Del.IrRmSz2    and SRmSz3= Del.IrRmSz3 and SRmStkRt= Del.IrStkRt and Del.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmDrQty=SRmDrQty-(select isnull(sum(Del.IrRmQty),0) from Deleted Del      
   where Del.IrToRmDc='D' and Del.IrPrtKey=SPrtKey and Del.IrCoCd= SCoCd and Del.IrRmCd=SRmCd and Del.IrLotNo=SLotNo       
   and SRmSz= Del.IrRmSz    and SRmSz2= Del.IrRmSz2    and SRmSz3= Del.IrRmSz3 and SRmStkRt= Del.IrStkRt and Del.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmDrWt=SRmDrWt-(select isnull(sum(Del.IrRmAWt),0) from Deleted Del      
    where Del.IrToRmDc='D' and Del.IrPrtKey=SPrtKey and Del.IrCoCd= SCoCd and Del.IrRmCd=SRmCd and Del.IrLotNo=SLotNo       
   and SRmSz= Del.IrRmSz    and SRmSz2= Del.IrRmSz2    and SRmSz3= Del.IrRmSz3 and SRmStkRt= Del.IrStkRt and Del.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm)      
  from Deleted       
  where Deleted.IrPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and Deleted.IrPrtKey=SPrtKey and Deleted.IrCoCd= SCoCd       
   and Deleted.IrRmCd= SRmCd and Deleted.IrLotNo= SLotNo and  SRmSz= Deleted.IrRmSz       
	and  SRmSz2= Deleted.IrRmSz2 and  SRmSz3= Deleted.IrRmSz3              
   and SRmStkRt= Deleted.IrStkRt and Deleted.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))= SYyMm      
   and (Deleted.IrToRmDc='D' or Deleted.IrToRmDc='C') and      
   ((@Mode='C' and Exists(Select 'x' From Inserted ins where ins.IrCoCd =Deleted.IrCoCd and ins.IrTc=Deleted.IrTc and ins.IrYy=Deleted.IrYy and      
         ins.IrChr=Deleted.IrChr and ins.IrNo=Deleted.IrNo and ins.IrSr=Deleted.IrSr and ins.IrSrNo=Deleted.IrSrNo and          
      (ins.IrPrtKey<>Deleted.IrPrtKey Or  ins.IrCoCd<>Deleted.IrCoCd Or ins.IrRmCd<>Deleted.IrRmCd Or ins.IrLotNo<>Deleted.IrLotNo      
        Or  ins.IrRmSz<>Deleted.IrRmSz Or  ins.IrRmSz2<>Deleted.IrRmSz2 Or  ins.IrRmSz3<>Deleted.IrRmSz3 or  ins.IrStkRt<>Deleted.IrStkRt Or ins.IrToRmLoc<>Deleted.IrToRmLoc Or       
Ins.IrRmAWt<>Deleted.IrRmAWt Or Ins.IrRmQty<>Deleted.IrRmQty Or ins.IrToRmDc<>Deleted.IrToRmDc Or    
         convert(int,convert(varchar(6),ins.IrDt,112))<>convert(int,convert(varchar(6),Deleted.IrDt,112))))) Or @Mode='D')      
      
      
         
 end      
      
 if @mode in ('A','C')      
 begin      
  /* Current Partition */      
  insert into SYyMm(SPrtKey,SCoCd, SRmCd,SLotNo,SRmSz,SRmSz2,SRmSz3, SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)       
   select distinct IrPrtKey,IrCoCd, IrRmCd,IrLotNo,IrRmSz,IrRmSz2, IrRmSz3,IrStkRt,IrToRmLoc,convert(int,convert(varchar(6),Indt,112)),0,0,0,0      
   from inserted join InvHd on InPrtKey=IrPrtKey and InCoCd=IrCoCd and InTc=IrTc and InYy=IrYy and InChr=IrChr and InNo=IrNo      
   where @wTcTyp= 'IN' and not exists(select 'x' from SYyMm where       
    SPrtKey=IrPrtKey and SCoCd= IrCoCd and SRmCd=IrRmCd and SLotNo=IrLotNo and SRmSz= IrRmSz       and SRmSz2= IrRmSz2 and SRmSz3= IrRmSz3              
    and SRmStkRt= IrStkRt  and SLoc= IrToRmLoc       
    and SYyMm=convert(int,convert(varchar(6),Indt,112)))       
       
  /* Current Partition */      
  update SYyMm set      
  SRmCrQty= SRmCrQty+(select isnull(sum(Ins.IrRmQty),0) from Inserted Ins      
   where Ins.IrToRmDc='C' and Ins.IrPRtKEy=SPrtKey and Ins.IrCoCd= SCoCd and Ins.IrRmCd=SRmCd and Ins.IrLotNo=SLotNo       
   and SRmSz= Ins.IrRmSz and SRmSz2= Ins.IrRmSz2 and SRmSz3= Ins.IrRmSz3 and SRmStkRt= Ins.IrStkRt and Ins.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmCrWt= SRmCrWt+(select isnull(sum(Ins.IrRmAWt),0) from Inserted Ins      
   where Ins.IrToRmDc='C' and Ins.IrPrtKey=SPrtKey and Ins.IrCoCd= SCoCd and Ins.IrRmCd=SRmCd and Ins.IrLotNo=SLotNo       
   and SRmSz= Ins.IrRmSz    and SRmSz2= Ins.IrRmSz2    and SRmSz3= Ins.IrRmSz3 and SRmStkRt= Ins.IrStkRt and Ins.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmDrQty= SRmDrQty+(select isnull(sum(Ins.IrRmQty),0) from Inserted Ins      
   where Ins.IrToRmDc='D' and Ins.IrPrtKey=SPrtKey and Ins.IrCoCd= SCoCd and Ins.IrRmCd=SRmCd and Ins.IrLotNo=SLotNo       
   and SRmSz= Ins.IrRmSz    and SRmSz2= Ins.IrRmSz2    and SRmSz3= Ins.IrRmSz3 and SRmStkRt= Ins.IrStkRt and Ins.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm),      
  SRmDrWt= SRmDrWt+(select isnull(sum(Ins.IrRmAWt),0) from Inserted Ins      
   where Ins.IrToRmDc='D' and Ins.IrPrtKey=SPrtKey and Ins.IrCoCd= SCoCd and Ins.IrRmCd=SRmCd and Ins.IrLotNo=SLotNo       
   and SRmSz= Ins.IrRmSz    and SRmSz2= Ins.IrRmSz2    and SRmSz3= Ins.IrRmSz3 and SRmStkRt= Ins.IrStkRt  and Ins.IrToRmLoc=SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm)      
  from Inserted      
  where Inserted.IrPrtKey=@wCurrPrtn and @wTcTyp= 'IN' and Inserted.IrPRtKey=SPrtKey and Inserted.IrCoCd= SCoCd       
   and Inserted.IrRmCd= SRmCd and Inserted.IrLotNo= SLotNo and SRmSz= Inserted.IrRmSz       
   and SRmSz2= Inserted.IrRmSz2 and SRmSz3= Inserted.IrRmSz3              
   and SRmStkRt= Inserted.IrStkRt and Inserted.IrToRmLoc= SLoc       
   and convert(int,convert(varchar(6),IrDt,112))=SYyMm      
   and (Inserted.IrToRmDc='D' or Inserted.IrToRmDc='C') and      
   ((@Mode='C' and Exists(Select 'x' From Deleted Del where inserted.IrCoCd =Del.IrCoCd and inserted.IrTc=Del.IrTc and inserted.IrYy=Del.IrYy and      
        inserted.IrChr=Del.IrChr and inserted.IrNo=Del.IrNo and inserted.IrSr=Del.IrSr and inserted.IrSrNo=Del.IrSrNo and          
     (inserted.IrPrtKey<>Del.IrPrtKey Or  inserted.IrCoCd<>Del.IrCoCd Or inserted.IrRmCd<>Del.IrRmCd Or inserted.IrLotNo<>Del.IrLotNo      
       Or  inserted.IrRmSz<>Del.IrRmSz Or  inserted.IrRmSz2<>Del.IrRmSz2 Or  inserted.IrRmSz3<>Del.IrRmSz3 or  inserted.IrStkRt<>Del.IrStkRt Or inserted.IrToRmLoc<>Del.IrToRmLoc Or       
       inserted.IrRmAWt<>Del.IrRmAWt Or inserted.IrRmQty<>Del.IrRmQty Or inserted.IrToRmDc<>Del.IrToRmDc Or       
        convert(int,convert(varchar(6),inserted.IrDt,112))<>convert(int,convert(varchar(6),Del.IrDt,112))))) Or @Mode='A')      
      
 end      
/*********************************** urmi 5/7/03 Ch28  *************************************/      
      
 return      
ErrHndlr:      
 begin   select @errmsg='|'+@errmsg   raiserror(@errmsg,16,1)   rollback transaction       
 end      
end 


/********************************* Loose Material *******************************/


end



GOGO


ALTER TRIGGER [dbo].[OrdDsg_INUPDL] ON [dbo].[OrdDsg]     
   FOR INSERT,UPDATE,DELETE     NOT FOR REPLICATION
As    
BEGIN    
 DECLARE @ROW INT    
 DECLARE @mode char(1)    
 DECLARE @errmsg varchar(255)    
 DECLARE @modUsr char(5)
 Declare @wSelf VarChar(8)    
 Declare @wSysDb Varchar(3)    
 DECLARE @ReplAgt BIT    
 DECLARE @wCurrPrtn VarChar(1)    
 Declare @wInFob Float  
 Declare @wInFobRs Float  
 Declare @wInLclOctroiRs Float  
 Declare @wInLclVatRs Float  
 Declare @wInLclCstRs Float  
 DECLARE @wTcTyp varchar(3)   
 declare @wIsJLTc char(1)
 IF @@ROWCOUNT=0   RETURN    
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'OrdDsg' )    
  RETURN    
    
 Select @mode= case (select count('x') from inserted) when 0 then 'D'    
        else    
    case (select count('x') from deleted) when 0 then 'A' else 'C' end    
        end     
    
 Select @wSysDb=SysDb From SysDb where SysId='CRM'     
 Select @wCurrPrtn = 'C'       
 Set @wSysDb=IsNull(@wSysDb,'')     
 declare @wModUsr varchar(5)
 select @wModUsr=''
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))    
 If @ReplAgt = 1 Return
 Select @wSelf='ZSELF'     
 if @mode in ('A','C')  
 begin  
  select @wModUsr=max(inserted.ModUsr), @wTcTyp= max(PValue),@wIsJLTc = max(PValue1) from inserted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''  
  if @wTcTyp = '' select @wTcTyp = 'N'

 end  
  
 if @mode in ('D')  
 begin  
  select @wTcTyp=max(PValue), @wIsJLTc = max(PValue1) from deleted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''   
 End  
 --5717
 /******** Only One TcTyp records can be Inserted at a time.********/   
 select @Row= Count(distinct(PValue)) from inserted Join param On  
  PTyp= 'TC' and PMCd= OdTc and PScd= '' Where OdPrtKey=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'  
  GoTo ErrHndlr  
 end  
 --5717 
 /******** Only One TcTyp records can be Deleted at a time. *********/   
 select @Row= Count(distinct(PValue)) from deleted Join param on  
  PTyp= 'TC' and PMcd= OdTc and PSCd='' Where OdPrtKey=@wCurrPrtn  
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted on'  
  GoTo ErrHndlr  
 end  
   
 Select @Row= Count('x') from deleted where OdTc = 'JRI' and exists 
 (Select 'a' from Txnd where TdCoCd = OdCoCd and TdBYy = OdDmStkYy and TdBChr = OdDmStkChr and TdBNo = OdDmStkNo and (Txnd.ModDt > deleted.ModDt or (Txnd.ModDt = deleted.ModDt and Txnd.ModTime > Deleted.ModTime)))
 If @Row > 0  
 begin   
  Select @ErrMsg= 'Cannot Modify/Delete As Bag Transactions exists for Repair Bag'  
  GoTo ErrHndlr  
 end

 if @mode in ('C','D') and @wTcTyp = 'JT'
 Begin
	 Select @row= count('x') from deleted del where 
	   not exists(Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdDmCd = del.OdDmCd 
	   and ins.OdSfx= del.OdSfx and ins.OdDmSz = del.OdDmSz and ins.OdDmStkCoCd = del.OdDmStkCoCd and ins.OdDmStkYy = del.OdDmStkYy
	   and ins.OdDmStkChr = del.OdDmStkChr and ins.OdDmStkNo= del.OdDmStkNo) and
	   exists (Select 'a' from OrdRm where OrOdMstIdNo = del.OdIdNo)
	 If @row > 0     
	 begin    
	  Select @ErrMsg= 'Cannot change DmCd, Sfx, Sz or DmStkNum or delete when this is the master record for JRO transactions'  
	  goto ErrHndlr    
	 End   
 End
     
  If @Mode ='D'     
  Begin    
   Delete OrdRm From deleted, OrdRm where OdIdNo=OrOdIdNo and OdCoCd = OrCoCd and OdTc = OrTc and     
    OdYy = OrYy and OdChr = OrChr and OdNo = OrNo and OdSr= OrSr 
     
   Delete OrdLab From deleted, OrdLab where OdIdNo=OlOdIdNo and OdCoCd = OlCoCd and OdTc = OlTc and     
    OdYy = OlYy and OdChr = OlChr and OdNo = OlNo and OdSr= OlSr 
     
   Delete OrdAmd From deleted, OrdAmd where OdIdNo=OaOdIdNo and OdCoCd = OaCoCd and OdTc = OaTc and     
    OdYy = OaYy and OdChr = OaChr and OdNo = OaNo and OdSr= OaSr 
     
   Delete OrdTgt From deleted, OrdTgt where OdCoCd = OtCoCd and OdTc = OtTc and     
    OdYy = OtYy and OdChr = OtChr and OdNo = OtNo and OdSr = OtSr 
     
   Delete OrdPln From deleted, OrdPln where OdCoCd = OpCoCd and OdTc = OpTc and     
    OdYy = OpYy and OdChr = OpChr and OdNo = OpNo and OdSr = OpSr     
    
   Delete OrdCT From deleted, OrdCT where OdIdNo=OctOdIdNo and OdCoCd = OctCoCd and OdTc = OctTc and     
    OdYy = OctYy and OdChr = OctChr and OdNo = OctNo and OdSr = OctSr 
    End      
 --End    
	  
	Select @row= count('x') from inserted ins where ins.OdHld= 'Y'     
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

	  If @mode = 'C'    and (update(OdMinDiatolwt) or update(odMaxdiatolWT))
	  Begin      
	 Select @row=count('x') from inserted Where OdMinDiaTolWt > OdMaxDiaTolWt  
	 If @row > 0  
	 Begin  
	  Select @Errmsg='Minimunm Diamond Tolerance cannot be Greater Than Maximum Diamond Tolerance'  
	  Goto ErrHndlr  
	 end   
    
	 
 End    
    
  if @mode = 'A'     
  Begin    
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo 
	
    where OmReYy= '' And OmReChr= '' And OmReNo=0 And @wIsJLTc = 'N' and
    
    not exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp     
        and DmHld= 'N'  and DmCd= OdDmCd and DmSz= ''  )    
  if @row> 0     
  begin    
   select @errmsg= 'Design/Sketch Code Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from inserted     
    where @wIsJLTc = 'N' and
    not exists (Select PValue from Param where PTyp= 'DMTCTYP' and PMCd= OdDMTcTyp)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Tc Type is Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end
    
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo 
	
   where OmReYy= '' And OmReChr= '' And OmReNo=0 And  @wIsJLTc = 'N' and    
	@wIsJLTC='N' and 
    Exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp     
     and (DmHld= 'Y' ) and DmCd= OdDmCd and DmSz= OdDmSz )    
  if @row> 0     
  begin    
   select @errmsg= 'This Design/Sketch Size Is On Hold/Invalid'    
   goto ErrHndlr    
  end  
  --5717  
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) <> 'QB' And  
  (Select PMCd from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT') <> 'sri')   
  and (OdBYy <> '' or OdBChr <> '' or OdBNo <> 0)    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Should Be Blank For This Voucher Type'    
   goto ErrHndlr    
  end 

  -- Check if Not exists part with OrdDsg makes sense
  select @row= count('x') from inserted Ins where 
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('QB'))   
 and not exists (Select 'x' From Bag Where BCoCd= OdCoCd and BYy= OdBYy     
      and BChr = OdBChr And BNo= OdBNo and BLoc = 'PFG'     
      and BOdDmCd= OdDmCd )    
  /* And Not Exists (Select 'x' From OrdDsg Od Where Od.OdCoCd=ins.OdCoCd And Od.OdTc In ('QB') And     
      Od.OdBYy=Ins.OdBYy And Od.OdBChr=Ins.OdBChr And Od.OdBNo=Ins.OdBNo And Od.OdPrtKey=Ins.OdPrtKey)  */  
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Not Defined For the Design'    
   goto ErrHndlr    
  end    
end 
 if @mode in ('A','C') and @wModUsr<>'MIC'     
 Begin    
  select @row= count('x') from inserted where not exists (select 'x' from OrdMst where     
     OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and     
     OmChr= OdChr and OmNo= OdNo )    
  if @row> 0     
  begin    
   select @errmsg= 'Voucher Not Defined In Order Master'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdSr<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Design Serial Number Cannot Be <= Zero'    
   goto ErrHndlr    
  end  
  --5717
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
    
  select @row= count('x') from inserted Ins where OdDmSz<> '' and     
   Ins.OdDmSz <> IsNull((Select Del.OdDmSz From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
    Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy     
    And Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr ), '') And     
   not exists (select 'x' from Param where PTyp= 'DMSZ' and PMCd= OdDmSz And PValidYn = 'Y')    
  if @row> 0     
  begin    
   select @errmsg= 'Design Size Not Defined / Invalid.'    
   goto ErrHndlr    
  end    

  Select @row= count('x') from inserted where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and not exists (select 'x' from vParam where vPCoCd= OdCoCd and     
     vPTyp= 'PRDSEQ' and vPMCd= OdPrdSeq )    
  if @row> 0     
  begin    
   select @errmsg= 'Production Sequence Not Defined'    
   goto ErrHndlr    
  end   
    
  select @row= count('x') from inserted where OdPrdQty<  OdFgQty     
  if @row> 0     
  begin    select @errmsg= 'Production Qty Should Be >= Finished Goods Qty'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and    
   not exists (select 'x' from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd )    
  if @row> 0     
  begin    select @errmsg= 'Priority Code Not Defined'    
   goto ErrHndlr    
  end   
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPRtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdHld)    
  if @row> 0     
  begin    
   select @errmsg= 'Enter On Hold (Y/N)'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where OdPRtKey=@wCurrPRtn and OdHld= 'N' and OdHldDesc<> ''     
  if @row> 0     
  begin    
   select @errmsg= 'Holding Desc Should Be Entered Only When Design Is On Hold'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdMulBy<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Multiplying Factor Cannot Be <= Zero'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdFixPrc)    
  if @row> 0     
  begin    select @errmsg= 'Enter Fix Price (Y/N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPRtn and OdGldAs<> '' and     
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdGldAs)           
  if @row> 0     
  begin    
   select @errmsg= 'Gold As Value Not Defined'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdLabAs)    
  if @row> 0     
  begin    
   select @errmsg= 'Labour As Value Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdIWtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. Equal To Order Wt. (Y/N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdIWtFrOrd<> '' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'IWTFRORD' and PMCd= OdIWtFrOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. From Order Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdIWtFrOrd<> '' and OdIWtEqOrd= 'Y')     
  if @row> 0     
  begin    
   select @errmsg= 'Mismatch in values InvWtEqOrd and InvWtFrOrd'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'INRTORD' and PMCd= OdIGldRtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invalid value for Gold Rate Equal To Order Rate'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and     
     not exists (select 'x' from Param where PTyp='YN' and PMCd= OdCls)    
  if @row> 0     
  begin    
   select @errmsg= 'Order Closed (Y/N)'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'WH' and PMCd= OdWh)    
  if @row> 0      
  begin    
   select @errmsg= 'W/H Set Option Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdWDiaAvlblDt in('01/01/1980', '01/01/1900') or     
   OdHDiaAvlblDt in('01/01/1980', '01/01/1900') or OdFndAvlblDt in('01/01/1980', '01/01/1900'))    
  if @row> 0     
  begin    
   select @errmsg= 'Wax/ Hand Set Diamond Availability/Fnd Availability Date Cannot Be Blank'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPRtn  and OdTc <> 'JRI' and OdTc <> 'JRO' and    
    not exists (select 'x' from Param where PTyp= 'CPYRT' and PMCd= OdGmChk)    
  if @row> 0     
  begin    
   select @errmsg= 'Gram Check Option Not Defined'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
    not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChkTol)    
  if @row> 0     
  begin    
   select @errmsg= 'Should Tolerence Be Checked? (Specify Y/ N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and     
    ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and     
    ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdDmCd= del.OdDmCd)     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design'    
   goto ErrHndlr    
  end  
  --5717
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
  --5717
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
   
  --5717
  select @Row= Count('x') from inserted join InvHd on InCoCd= OdCoCd and InTc= OdTc and   
   InYy= OdYy and InChr= OdChr and InNo= OdNo where InCls= 'Y' and OdTc='FOT'  
  If @Row > 0  
  begin   
	Select @ErrMsg= 'Cannot Add/Change As Invoice Is Closed'  
	GoTo ErrHndlr  
  end  
  
    
  /******  OdVaCtg  =  PMCd ('VACTG')  ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where 
   not exists (select 'x' from Param where PTyp= 'VACTG' and PMCd= OdVaCtg)    
  if @row> 0     
  begin    
   select @errmsg= 'Value Addition Category Not Defined'    goto ErrHndlr    
  end    
    
  /******  '  =  PMCd ('KT')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and    
   not exists (select 'x' from Param where PTyp= 'KT' and PMCd= OdKt)    
  if @row> 0     
  begin    
   select @errmsg= 'Karatage Not Defined'    
   goto ErrHndlr    
  end    
    
  Select @row= count('x') from inserted Ins where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and    
   (Ins.OdDmCol= '' or     
   ( Ins.OdDmCol <> IsNull((Select Del.OdDmCol From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
      Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy And     
      Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr ), '') And     
    not exists (select 'x' from Param where PTyp= 'DMCOL' and PMCd= OdDmCol And PValidYn = 'Y')     
   ))     
  if @row> 0     
  begin     
   select @errmsg= 'Design Colour Not Defined / Invalid.'     
   goto ErrHndlr     
  End     

  Select @row= count('x') from inserted where @wIsJLTc = 'N' and 
   not exists (select 'x' from vParam where vPCoCd= OdCoCd and vPTyp= 'LINE' and vPMCd= OdLine and vPMCd<> 'ZZZ')    
  if @row> 0    begin    
   select @errmsg= 'Line Not Defined/ Line Cannot Be ''ZZZ'' '    
   goto ErrHndlr    
  end    
    
  
  select @row= count('x') from  deleted del where 
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and     
     ins.OdDmSz= del.OdDmSz and ins.OdSfx= del.OdSfx and ins.OdOrdEnt= del.OdOrdEnt and ins.OdPrdEnt= del.OdPrdEnt    
      and ins.OdPrdSeq= del.OdPrdSeq) 
	     and (exists(Select 'x' from Bag where BOdIdNo=del.OdIdNo))    


  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design Size/ Suffix/ OrdEnt/ PrdEnt / PrdSeq When Bag Exists'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted ins     
 join OrdMst insOm On insOm.OmIdNo=ins.OdOmIdNo       
  where 
     (Select PValue from Param where PTyp='TC' and PMCd= ins.OdTc)= 'PL' and     
     (Select count('x') from OrdDsg Od join OrdMst Om on  Om.OmIdNo=Od.OdOmIdNo where 
		Om.OmCoCd= insOm.OmCoCd and Om.OmTc= (Case When insOm.OmCmCtg = 'S' Then 'SPL' Else 'PL' End)     
    and Om.OmCmCtg= insOm.OmCmCtg and Om.OmCmCd= insOm.OmCmCd and Od.OdDmCd= ins.OdDmCd     
    and Od.OdSfx= ins.OdSfx and Od.OdDmSz= ins.OdDmSz and Om.OmCmCurCd= insOm.OmCmCurCd) > 1     
  if @row> 0     
   begin    
     select @errmsg= 'Price List Already Defined for this Cust, Curr, Dsg, Sz, Sfx combination'    
     goto ErrHndlr    
   End    

  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdDelDt= '01/01/1900' or (OdDelDt<> '01/01/1980' and     
   OdDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc     
     and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo))     
  if @row> 0     
  begin    
   select @errmsg= 'Production Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr     
  end     

  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdExpDelDt= '01/01/1900' or (OdExpDelDt<> '01/01/1980' and     
   OdExpDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd     
    and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo))     
    
  if @row> 0     
  begin    
   select @errmsg= 'Export Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr   end    
    
  Select @row= count('x') from inserted where oDtC <> 'JCN' and @wIsJLTc = 'N' and OdOrdQty < 0 or OdPrdQty< 0    
  if @row> 0     
  begin    select @errmsg= 'Order/ Prd Qty Should Be >= 0 '    
   goto ErrHndlr    
  end    
    
 select @row= count('x') from inserted where odtc <> 'JCN' and @wIsJLTc = 'N' and 
   Round(OdOrdQty, 1) < Round(OdExpQty, 1)     
  
  if @row> 0     
  begin    select @errmsg= 'Order Qty Should Be >= Export Qty '    
   goto ErrHndlr    
  end    
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
  --5717
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
  --5717
   select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdBagPcs<= 0 and OdTc <> 'JRI' and OdTc <> 'JRO'        
  if @row> 0     
  begin    
   select @errmsg= 'Default Bag Opening Pieces Cannot Be <= Zero'    
   goto ErrHndlr    
  end 

  select @row= count('x') from inserted where OdOmCtCd<> '' and     
   not exists (select 'x' from CustMst where CmCtg= 'T' and CmCd= OdOmCtCd)    
  if @row> 0     
  begin     
   select @errmsg= 'Contractor Code Not Defined'    
   goto ErrHndlr    
  end    
    
  
  select @row= count('x') from inserted ins     
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo 
  
  where OdGrMet<>''     
   and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'MET' and PSCd= OdGrMet)     
  if @row> 0     
  begin     
   select @errmsg= 'Metal Grade not Defined'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted ins    
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo  
  where OdGrDia<>'' and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'DIA' and     
   PSCd= OdGrDia)     
  if @row> 0     
  begin     
   select @errmsg= 'Diamond Grade not Defined'    
   goto ErrHndlr    
  end    
    
  if @wTcTyp = 'JT'   
	  Begin   
	  -- rewrite 

	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmIdNo=Ins.OdOmIdNo 
	Where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
			Where Od.OdCoCd=(Case when ins.OdTc in ('JIN','LIN')  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
			
	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmIdNo=Ins.OdOmIdNo 
	Where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       ((TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
			
			Where Od.OdCoCd=Ins.OdCoCd  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo Or Od.OdDmStkNo=0)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))  
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
   
   Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo 
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
          
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
            Where Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1 =Ins.OdJLLn1 and
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.OdJLLn3    
            and (Od.OdDmStkNo=Ins.OdDmStkNo/*Or (Od.OdDmStkNo=0 Or Ins.OdDmStkNo=0)*/)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))
			
	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo 
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
	   (
       (TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
            Where Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and (Od.OdJLLotNo=Ins.OdJLLotNo or Od.OdJLLotNo = '') and Od.OdJlLn1 = Ins.OdJlLn1 and 
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))  
       )  
        )  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
   Select @Row=Count('x') from Inserted Ins Where 
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
           ), '')+',%')  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  
	Select @Row=Count('x') from Inserted Ins where
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.ODJLRmCd =Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1=Ins.OdJLLn1 
				and Od.ODJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
           ), '')+',%')  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  

	Select @row= count('x') from inserted where 
    exists(Select 'x' from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT') and Round(OdPrdQty, 1) > 0    
	   	   	   	   if @row> 0     
    begin    select @errmsg= 'Prd Qty should be 0 for Trading voucher'    
     goto ErrHndlr    
    end    

     
  end  
  

End    
--5717
if @mode = 'C'    
	  Begin      
	  select @Row= Count('x') from deleted del where     
		not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
		and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
		and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr)    
	  If @Row > 0    
	  begin     
	   Select @ErrMsg= 'Cannot Change Company Code/Tc/Yy/Chr/No/Sr'    
	   GoTo ErrHndlr    
	  end    
	 --5717
	  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and   
	  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChgPtrOnStwYN and PSCd= '')  
	 if @row> 0     
	 begin  
	  select @errmsg= 'Change Pointer On STW (Y/N)'  
	  goto ErrHndlr  
	 end   
	 --5717
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
	 --5717
	 select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and  
	  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdILabWtFrOrd and PSCd= '')  
	 if @row> 0     
	 begin  
	  select @errmsg= 'Labor Weight From Order (Y/N)'  
	  goto ErrHndlr  
	 end  
	 --5717
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
end    

 if @mode = 'D'   
 Begin    
  Select @row= count('x') from deleted where 
   (exists(Select 'x' from Bag where     
                BOdIdNo=OdIdNo ))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Export Has Taken Place / Design Under Production'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and      	 
    exists(Select 'x' from OrdLn where OdlCoCd =OdCoCd and OdlTc =OdTc and OdlYy =OdYy and OdlChr =OdChr 
    and OdlNo =OdNo and OdlSr =OdSr and OdlOdIdNo =OdIdNo )   
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Transaction Exist in OrdLn'    
   goto ErrHndlr    
  end   
  --5717
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and    
   (OdPrtExpQty<>0 Or exists(Select 'x' from Fgd where FdCoCd= OdCoCd and FdExpOdTc= OdTc and     
         FdExpOdYy = OdYy and FdExpOdChr = OdChr and FdExpOdNo = OdNo and     
         FdExpOdSr = OdSr and FdPrtKey=OdPrtKey))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Record Exists As Export Order In Fgd Table'    
   goto ErrHndlr    
  end 
    
  if @wTcTyp = 'JT'   
  Begin   
  
  /* Cannot Delete JRI if exist in Txnd */
  select @row= count('x') from deleted del where del.OdTc='JRI' and del.OdDmStkYy <> ''  and del.OdDmStkChr <> '' and del.OdDmStkNo <> ''  
    and exists(Select 'x' from Txnd tx where tx.TdCoCd =del.OdCoCd and tx.TdBYy = del.OdDmStkYy and tx.TdBChr = del.OdDmStkChr and 
		tx.TdBNo = del.OdDmStkNo and tx.ModDt >= del.ModDt and tx.ModTime >= del.ModTime)       
  if @row > 0    
  begin    
   select @errmsg= 'Cannot Delete, Bag Transaction Exists for this Repair Bag'    
   goto ErrHndlr    
  end      
  /* why is OdBaseIdKey required to be checked index is there or not */
  Select @row=Count('x') from deleted Del Where 
    exists(Select 'x' from OrdDsg Od Where   
       (Od.OdBaseIdKey=Del.OdIdNo))  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Base Voucher'  
    goto ErrHndlr  
   End  
    
   Select @row=Count('x') from deleted Del Where 
    exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and 
       Od.OdRefIdKey=Del.OdIdNo)  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Reference Voucher'  
    goto ErrHndlr  
   End  
  End   
 End    
   
 /* **** Manali Trading Module **** */  
 if @mode ='C'  
 Begin  
 if @wTcTyp = 'JT'   
  Begin  
  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where @wIsJLTc = 'N' and
   exists(Select 'x' from OrdDsg Od Where 
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
    and ins.OdDmCd=del.OdDmCd and ins.OdSfx=del.OdSfx and ins.OdDmSz=del.OdDmSz   
	and ins.OdDmStkYy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr 
    and ins.OdDmStkNo=del.OdDmStkNo
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  
      
  Select @row=Count('x') from deleted Del Where 
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and 
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'N' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and 
   ((ins.OdDmStkyy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr and ins.oddmstkno = del.oddmstkno and ins.OdDmStkNo=del.OdDmStkNo) or ins.OdDmStkNo=0 or del.OdDmStkNo=0))    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  

 End  

  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where @wIsJLTc = 'Y' and
   exists(Select 'x' from OrdDsg Od Where  (Del.OdTc='LIN' Or  Od.OdCoCd=Del.OdCoCd) and   
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
    and ins.OdJLRmCd=del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJLLn2 = del.odJLLn2 and ins.odJLLn3 = del.odJLLn3 
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  

    
  Select @row=Count('x') from deleted Del Where   
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and   
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'Y' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
     and ins.OdJLRmCd =del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJlLn2 = del.OdJlLn2 and ins.OdJLLn3 = del.ODJLLn3
     )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  
 End  

 --??End  
  
 if @mode in ('C','D')  
 begin  
 If @wTcTyp = 'FG' Or @wTcTyp = 'JT'         
 BEgin  
  update InvHd set   
  @wInFob= convert (decimal (16, 2), InFOB- IsNull((Select sum((del.OdSalPrc*del.OdOrdQty)+del.OdJLVchVal) from deleted del where del.OdCoCd= InCoCd and del.OdTc= InTc and del.OdYy= InYy   
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
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo   
  where InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD In ('FG', 'JT') and PValue= OdTc)  
         or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  
  if @wIsJLTc = 'N'
  begin
  
  update DmYyMm set  
  DDmCrQty= DDmCrQty- (select isnull(sum(Del.OdOrdQty), 0) from Deleted Del  
   Join OrdMst on   
   OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm),  
  DDmDrQty= DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkyy = DDmStkYy and del.OdDmStkChr = 
			DDmStkChr and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd    
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkCoCd =DDmStkCoCd and Deleted.OdDmStkNo= DDmStkNo and Deleted.OdDmStkYy= DDmStkYy and Deleted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
   else
   begin
  update DmYyMm set  
    DDmCrWt= DDmCrWt- (select isnull(sum(Del.OdJLWt), 0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm),  
  DDmDrWt= DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd= DJLRmCd 
   and Deleted.OdJLLotNo= DJLLotNo and Deleted.OdJLLn1= DJLLn1 and Deleted.OdJLLn2= DJLLn2 and Deleted.OdJLLn3 = DJLLn3 and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
 End  
  end
    
	if @mode in ('A','C')
	begin
		Select @row= count('x') from inserted where Inserted.OdTc='JRI' and 
			exists (Select 'a' from Bag where BCoCd=OdDmStkCoCd and BYy=OdDmStkYy and BChr = OdDmStkChr and BNo=OdDmStkNo) and
				not exists 
				(Select 'a' from Bag where 
							BCoCd = OdDmStkCoCd and BOdDmCd=OdDmCd and BOdSfx=OdSfx and BOdDmTcTyp = 'DM' and
							BOdDmSz=OdDmSz and BYy = OdDmStkYy and BChr =OdDmStkChr and BNo = OdDmStkNo and BCls='Y' and BLoc='PFG' and BRepYN='N')
 
		if @row>0 
		begin
			select @errMsg = 'Given DmStkNo not a valid Bag No For this Dsg Code-Sz-Suffix or Bag not closed or Bag not in FG or Bag already in Repair'
			goto ErrHndlr
		end
	end 
 
 If @mode in ('A','C')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'   
 Begin  

  /*** Update InvHd ***/  
  update InvHd set  
  @wInFob = convert (decimal (16, 2), InFOB+ IsNull((Select sum((ins.OdSalPrc*ins.OdOrdQty)+ins.OdJLVchVal) from inserted ins where   
   ins.OdCoCd= InCoCd and ins.OdTc= InTc and ins.OdYy= InYy   
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
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo   
  where InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD ='FG' and PValue= OdTc)  
        or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  

  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr, DDmStkNo, DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt,DdmStkCoCd)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy,OdDmStkChr,OdDmStkNo,OmFrDmLoc,convert(int,convert(varchar(6),OdOmDt,112)),0,0,@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0,OdDmStkCoCd   
   from inserted join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DCoCd= OdCoCd and DDmCd=OdDmCd   
        and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy= DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmFrDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and OdDmStkCoCd =DdmStkCoCd and 
        DYyMm= convert(int,convert(varchar(6),OdOmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
         
  /* Current Partition */   
  if @wIsJLTc ='N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd =DDmStkCoCd and Ins.OdDmStkNo = DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd =DDmStkCoCd and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdDmCd= DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkCoCd =DDmStkCoCd  
   and Inserted.OdDmStkYy = DDmStkYy and Inserted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdCoCd= DCoCd 
   and OmFrDmLoc=DLoc  and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd= DJLRmCd 
   and Inserted.OdJLLotNo= DJLLotNo and Inserted.OdJLLn1= DJLLn1 and Inserted.OdJLLn2= DJLLn2 and Inserted.OdJLLn3 = DJLLn3  and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 end  
  end
  
 if @mode in ('C','D')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'     
 Begin  
 if @wIsJLTc = 'N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst  
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst   
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd   
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkCoCd =DDmStkCoCd and Deleted.OdDmStkNo = DDmStkNo and Deleted.OdDmStkYy = DDmStkYy and 
   Deleted.OdDmStkChr = OdDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
  else
  update DmYyMm set  
  DDmCrWt=DDmCrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst  
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst   
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd = DJLRmCd and Deleted.OdJLLotNo = DJLLotNo and Deleted.OdJLLn1 = DJLLn1 and
   Deleted.OdJlLn2 = DJLLn2 and Deleted.OdJllN3 = DJLLn3  and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
 If @wTcTyp = 'JT'     
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty- IsNull((Select sum(del.OdOrdQty) from deleted del where   
     (del.OdTc='JIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt- IsNull((Select sum(del.OdJLWt) from deleted del where   
    (del.OdTc='LIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from deleted where   
    OrdDsg.OdCoCd= deleted.OdCoCd and OrdDsg.OdIdNo=deleted.OdBaseIdkey  
   and Exists (Select 'x' from TcMst Where TmCoCd=deleted.OdCoCd and TmTc=deleted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls = (case when @wIsJLTc ='N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
				else (case when OrdDsg.OdJLWt= OrdDsg.OdJLExpWt and OrdDsg.OdJLWt<> 0 then 'Y' else 'N' end)   end)
   from deleted del where OrdDsg.OdCoCd= Del.OdCoCd and OrdDsg.OdIdNo= Del.OdBaseIdKey  
      
   and Exists (Select 'x' from TcMst Where TmCoCd=del.OdCoCd and TmTc=del.OdTC and TmBaseTCs<>'')  
 
  If @mode in ('D','C')
  Begin  
   Update Bag set
		BRepYN = 'N',   
		BRepQty = 0,
		BRepINDt                ='01/01/1980',
		BRepLoc  = '',
		BRepGrWt	= 0,
		BRepOdTc = '',
		BRepOdYy = '',
		BRepOdChr = '',
		BRepOdNo = 0,
		BRepOdSr = 0
   from Deleted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRI'

   Update Bag set
   		BRepYN = 'Y'
   from Deleted
		where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRO'

  End    

 End   
  
 End   
  
 if @mode in ('A','C')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'    
 Begin  
  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr,DDmStkNo,DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt,DdmStkCoCd)
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy, OdDmStkChr, OdDmStkNo,OmToDmLoc,convert(int,convert(varchar(6),OdOmDt,112)),0,0,
					@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0,OdDmStkCoCd  
     
   from inserted join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DCoCd= OdCoCd and DDmCd=OdDmCd   
       and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy = DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmToDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and OdDmStkCoCd =DdmStkCoCd and 
       DYyMm= convert(int,convert(varchar(6),OdOmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  
  if @wIsJLTc = 'N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd= DDmStkCoCd and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd= DDmStkCoCd and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdDmCd=DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkCoCd= DDmStkCoCd and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkChr = DDmStkChr and Inserted.OdDmStkYy = DDmStkYy 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd = DJLRmCd and Inserted.ODJLLotNo = DJLLotNo and Inserted.ODJLLn1 = DJLLn1 and Inserted.ODJLLn2 = DJLLn2 and 
   Inserted.ODJLLn3 = DJLLn3 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 If @wTcTyp = 'JT'   
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty + IsNull((Select sum(Ins.OdOrdQty) from Inserted Ins where   
   (ins.OdTc='JIN' Or Ins.OdCoCd= OrdDsg.OdCoCd)  and Ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt+ IsNull((Select sum(ins.OdJLWt) from Inserted Ins where   
   (ins.OdTc='LIN' Or ins.OdCoCd= OrdDsg.OdCoCd) and ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from Inserted where 
    OrdDsg.OdCoCd= Inserted.OdCoCd and OrdDsg.OdIdNo=Inserted.OdBaseIdkey  
    and Exists (Select 'x' from TcMst Where TmCoCd=Inserted.OdCoCd and TmTc=Inserted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls =(case when @wIsJLTc = 'N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
			else (case when OrdDsg.OdJLWt = OrdDsg.ODJLExpWt and OrdDsg.ODJLWt <> 0 then 'Y' else 'N' end) end)
   from Inserted Ins where (ins.OdTc = 'LIN' or ins.OdTc='JIN' Or  OrdDsg.OdCoCd= Ins.OdCoCd)   
   and OrdDsg.OdIdNo= Ins.OdBaseIdKey and Exists (Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTC and TmBaseTCs<>'')  
  if @mode in ('A','C')  
  Begin   
   Update Bag set   
		BRepYN = 'Y',
		BRepQty = OdOrdEnt,
		BRepINDt = getdate(),
		/* (Select OmDt From OrdMst Where OmCoCd= OdCoCd And OmTc = OdTc   
		And OmYy = OdYy And OmChr= OdChr And OmNo= OdNo And OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo)*/
		BRepLoc  = (Select HDefRepBagLoc from Head where HCoCd = Inserted.OdCoCd),
		BRepGrWt = OdRepGrWt,
		BRepOdTc = OdTc,
		BRepOdYy = OdYy,
		BRepOdChr = OdChr,
		BRepOdNo = OdNo,
		BRepOdSr = OdSr
   from Inserted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRI'

   Update Bag set
		BRepYN= 'N'
   from Inserted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRO'
  End  
 End  
  
 End    
  
  
if @mode in ('C','A') and @wTcTyp = 'JT'   
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') from Inserted Ins   
   where exists(Select 'x' from OrdDsg B Where  
      B.OdIdNo=Ins.OdBaseIdkey and B.OdOrdQty < B.OdExpQty)  
    and (exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>''))  
  If @Row> 0  
  begin   
   select @errmsg= 'Order Qty Should Be >= Matched Qty for Base Voucher'    
   goto ErrHndlr    
  end
 end
else
begin
 Select @Row=Count('x') from Inserted Ins   
   where 
    exists(Select 'x' from OrdDsg B Where  
      B.OdIdNo=Ins.OdBaseIdkey and cast(B.OdJLWt as decimal(10,3))< cast(B.OdJLExpWt as decimal(10,3)))  
    and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>'')  
 If @Row> 0  
  begin   
   select @errmsg= 'Order Wt Should Be >= Matched Wt for Base Voucher'    
   goto ErrHndlr    
  end
 end   
end    
  
If @mode in ('C','A') and  (@wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then 
	  (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) 
	  else (case when OmFrDmDc = 'C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Ins.OdCoCd= DCoCd and Ins.OdDmCd= DDmCd and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and   
    Ins.OdDmStkCoCd =DdmStkCoCd and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkNo = DDmStkNo and 
     (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Desing Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
end
else
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrWt)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.OdJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3=DJLLn3 and
	(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Size'  
     GoTo ErrHndlr  
     end  

end
end  
If @mode in ('D') and  (@wTcTyp = 'FG' or @wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 -- rewrite checking JCN why and remove FG and see what to simplify 
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd = Del.OdCoCd and LocTyp =(Case PValue When 'JT' Then 'S' End)  and   
            LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then (case when OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and   
    Del.OdDmStkCoCd =DdmStkCoCd and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo and 
    (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Design Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
	end
else
begin
 -- rewrite checking JCN why and remove FG and see what to simplify 
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Del.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End)  and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrwT)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Del.OdCoCd= DCoCd and Del.ODJLRmCd = DJLRmCd and Del.ODJLLotNo = DJLLotNo and Del.ODJLLn1 = DJLLn1 and Del.ODJLLn2 = DJLLn2 and Del.ODJLLn3 = DJLLn3 and
(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Sz'  
     GoTo ErrHndlr  
     end  
end

end  
return    
ErrHndlr:    
 begin   select @errmsg='|'+@errmsg    
  raiserror(@errmsg,16,1)   rollback transaction     
 end    
end   



GOGO


ALTER TRIGGER [dbo].[OrdLab_INUPDL] ON [dbo].[OrdLab] 
       FOR INSERT,UPDATE,DELETE  
		NOT FOR REPLICATION
AS  
BEGIN  
 /* Trigger Version 3.02.0 (Speed Optimization) */	
 DECLARE @ROW INT  
 DECLARE @mode char(1)  
 DECLARE @errmsg varchar(255)  
 Declare @wSysDb Varchar(3)
 DECLARE @ReplAgt BIT  
 DECLARE @wCurrPrtn VarChar(1)	

 IF @@ROWCOUNT=0   RETURN  
  
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'OrdLab' )  
  RETURN  
  
 select @mode=case (select count('x') from inserted) when 0 then 'D'  
       else  
   case (select count('x') from deleted) when 0 then 'A' else 'C' end  
       end   

	Select @wCurrPrtn='C'

	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))
	Select @wSysDb=SysDb From SysDb where SysId='CRM'
	Set @wSysDb=IsNull(@wSysDb,'')



 	/******** Key Cannot Be Changed ********/
	--5717
	if @mode = 'C'
 	Begin		
		select @Row= Count('x') from deleted del where del.OlPRtKey=@wCurrPrtn and
			 not exists (Select 'x' from inserted ins where ins.OlIdNo=del.OlIdNo and ins.OlCoCd= del.OlCoCd and ins.OlTc= del.OlTc
					and ins.OlYy= del.OlYy and ins.OlChr= del.OlChr and ins.OlNo= del.OlNo
					and ins.OlSr= del.OlSr and ins.OlSrNo= del.OlSrNo and ins.OlPrtKey=del.OlPrtKey)
		If @Row > 0
		begin
			Select @ErrMsg= 'Cannot Change Company Code / Tc/Yy/Chr/No/Sr/SrNo'
			GoTo ErrHndlr
		end
	End

	if @mode in ('A','C')
	begin
		/****** OrderSr should exist in OrdDsg table ******/
		/* Current Partition */
		select @row= count('x') from inserted where OlPrtKEy=@wCurrPrtn and
			not exists (select 'x' from OrdDsg where OdIdNo=OlOdIdNo and OdCoCd= OlCoCd
			and OdTc= OlTc and OdYy= OlYy and OdChr= OlChr
			and OdNo= OlNo and OdSr= OlSr and OdPrtKey=OlPrtKey)
		if @row> 0
		Begin
			select @errmsg= 'Voucher Sr Not Defined In Order Design'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where OlPrtKey=@wCurrPrtn and OlSrNo<=0
		if @row> 0
		Begin
			select @errmsg= 'Labour Serial Number Cannot Be <= Zero'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where OlPrtKey=@wCurrPrtn and
			not exists (select 'x' from Param where PTyp= 'LABMCD' and PMCd= OlMCd)
		if @row> 0
		Begin
			select @errmsg= 'Labour Main Code Not Defined'
			goto ErrHndlr
		end
	
		select @row= count('x') from Inserted Ins where Ins.OlPrtKEy=@wCurrPRtn and
			Ins.OlSCd <> isNull((Select Del.OlSCd from Deleted Del where ins.OlIdNo=del.OlIdNo and ins.OlCoCd= del.OlCoCd
				and ins.OlTc= del.OlTc and ins.OlYy= del.OlYy and ins.OlChr= del.OlChr
				and ins.OlNo= del.OlNo and ins.OlSr= del.OlSr and ins.OlSrNo= del.OlSrNo and ins.OlPrtKey=del.OlPrtKey), '') And
			not exists (select 'x' from Param where PTyp= 'LABSCD'
			and PMCd= OlMCd and PSCd= OlSCd And PValidYN = 'Y')
		if @row> 0
		Begin
			select @errmsg= 'Labour Sub Code Not Defined / InValid'
			goto ErrHndlr
		end
		select @row= count('x') from inserted where OlPrtKEy=@wCurrPrtn and
			not exists (select 'x' from Param where
				    PTyp= 'LABQW' and PMCd= OlQw and PSCd= '')
		if @row> 0
		begin
			select @errmsg= 'Mention If Labour Value Calculation Is By Quantity Or Weight'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where OlPrtKey=@wCurrPRtn and
			not exists (select 'x' from Param where PTyp= 'LABQW' and PMCd= OlCstQw and PSCd= '')
		if @row> 0
		begin
			select @errmsg= 'Mention If Labour Cost Value Calculation Is By Quantity Or Weight'
			goto ErrHndlr
		end

		/* Current Partition */
		select @row= count('x') from inserted where OlPrtKey=@wCurrPrtn and OlQty<= 0
		if @row> 0
		begin
			select @errmsg= 'Labour Quantity Should Be Greater Than Zero'
			goto ErrHndlr
		end
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


ALTER TRIGGER [dbo].[Txnd_INUPDL] ON [dbo].[Txnd]              
       FOR INSERT,UPDATE,DELETE NOT FOR REPLICATION
AS              
BEGIN              
	Declare @Row Int              
	Declare @Mode Char(1)              
	Declare @ErrMsg Varchar(255)              
	Declare @wTcTyp Char(3)               
	Declare @wTc Char(3)               
	Declare @wEntTyp Char(8)              
	Declare @wHCd Varchar(8)              
	DECLARE @ReplAgt BIT                
	DECLARE @wCurrPrtn VarChar(1)         
	Declare @wTChr Char(1) 
	Declare @AckYN Char(1)  
 	Declare @wSysDbUni Char(1)
	Declare @whOLORYN char(1)
	If @@RowCount=0              
	Return              
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'Txnd' )              
	RETURN              
              
	Select @Mode=Case (Select Count('x') From Inserted)              
	When 0 Then 'D'              
	Else Case (Select Count('x') From Deleted)              
	When 0 Then 'A'              
	Else 'C'              
		End              
	End              
              
	Select @wHCd= 'ZSELF'              
	Select @wCurrPrtn='C'              
      
	Select @wSysDbUni = 'N'
      
	Select @whOLORYN = hOLORYN from head where HCoCd ='ZZZ'
  
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))              

	/*Update 
		Bag 
	set 
		BLabRem='O:'+ISNULL(Del.TdFrBLoc,'')+':N:'+Ins.TdToBLoc+convert(varchar,Ins.TdIdNo)+':'+convert(varchar, getdate(), 120)
	from 
		Inserted Ins
	LEFT OUTER join 
		Deleted Del on Del.TdIdNo=Ins.TdIdNo
	where 
		(ins.TdTc='BV' or ins.TdTc='DT') and 
		(Ins.TdFrBLoc='' and Ins.TdBNo<>0) 
		and Ins.TdSrNo=0
		and Ins.TdBIdNo=Bag.BIdNo
		*/              
	If @Mode in ('A','C')              
	Begin              
	Select @wTcTyp= (PValue) From Inserted Join Param On PTyp='TC'              
		And PMcd= TdTc And PScd=''              
	Select @wTc= TdTc From Inserted Join Param On PTyp='TC'              
		And PMcd= TdTc And PSCd=''              
	Select @wTChr =  (vPValue3) from inserted join vParam on vPCoCd = TdCoCd and vPTyp='CHR'
		And vPMCd = TdTc and vPSCd = TdChr 
	End              
  
               
	If @Mode in ('D')              
	Begin              
	Select @wTcTyp= (PValue) From Deleted Join Param On PTyp='TC'              
		And PMcd=TdTc And PScd=''              
	End              
               
	select @wEntTyp = PValue1 from Param where PTyp='TCTYP'              
     and PMcd=@wTcTyp and PSCd=''              
                
	select @Row= Count('x') from Inserted Ins where Ins.TdSrNo<>0 and
	not exists (select count('x') from Txnd ForChk where ForChk.TdCoCd=Ins.TdCoCd
	and ForChk.TdTc=Ins.TdTc and ForChk.TdYy=Ins.TdYy and ForChk.TdChr = Ins.TdChr and ForChk.TdNo=Ins.TdNo
	and ForChk.TdSr=Ins.TdSr and ForChk.TdSrNo=0)
	If @Row > 0              
	begin              
		Select @ErrMsg= 'Data Issue please contact 9820490905'              
			GoTo ErrHndlr              
	end   

	If @wEntTyp= 'TxnBag'              
	BEGIN        
		if @mode in ('C') 
		begin 
			select @Row= Count('x') from Inserted              where
			not exists (select count('x') from deleted where inserted.TdBIdNo<>deleted.TdBIdNo)
		    If @Row > 0              
			begin              
		      Select @ErrMsg= 'Cannot change Bag number'              
			      GoTo ErrHndlr              
			end   
		end
	--5717
	select @Row= Count('x') from inserted              
    join Bag on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo               
      and TdPrtKey=BPrtKey              
    where BPrtKey=@wCurrPrtn and BLoc= 'PFG' and BRepYN <> 'Y'              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Add/ Change As Bag In Finished Goods Location'              
    GoTo ErrHndlr              
   end  
   --5717
   select @Row= Count('x') from deleted              
    join Bag on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo               
     and TdPRtKey=BPrtKEy              
    where BPrtKey=@wCurrPrtn and BLoc= 'PFG' and BRepYN <> 'Y'              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Change/ Delete As Bag In Finished Goods Location'              
    GoTo ErrHndlr              
   end     
   --5717
   select @Row= Count('x') from inserted where @wTcTyp= 'BV' and TdPRtKey=@wCurrPrtn and              
    (TdFrBLoc= 'PFG' or TdToBLoc= 'PFG')              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'From/ To Bag Location Cannot Be ''PFG'' In Bag Movement'              
    GoTo ErrHndlr              
   end    

   if @mode in ('A','C') 
   begin select @Row= Count('x') from Inserted              
    join Bag on TdBIdNo=BIdNo 
    where BLoc= 'PFG' and BRepYN = 'N'              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot enter, this Bag is In Finished Goods Location'              
    GoTo ErrHndlr              
   end   
   end
   
   if @mode in ('D')           
   begin
   Select @Row= Count('x') from deleted              
    join Bag on TdBIdNo=BIdNo where BLoc= 'PFG' and BRepYN = 'N'              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Delete As Bag In Finished Goods Location'              
    GoTo ErrHndlr              
   end 
   end             

   if @mode in ('A','C')
   begin 
	select @Row= Count('x') from inserted              
    join Bag on TdBIdNo=BIdNo           
    where Not ((TdCoCd= BCoCd and TdPrevYy= BLstYy and TdPrevKey= BLstKey)              
    Or (TdCoCd= BCoCd and TdYy= BLstYy and TdKey= BLstKey))              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Bag in Control of other user'              
    GoTo ErrHndlr              
   end              
   end
	
   if @mode = 'D'
   begin
   select @Row=Count('x') from deleted where 
    Not Exists (Select 'x' from Bag where TdBIdNo=BIdNo  and              
         ((TdYy>= BLstYy and TdKey>= BLstKey)              
           or IsNull(BLstKey, 0)= 0))              
   If @Row > 0              
   Begin              
    Select @ErrMsg= 'Cannot Update Or Delete, Not The Last Txn Of The Bag'              
    GoTo ErrHndlr              
   end  
   end 
              
  If @mode in ('A')              
  Begin
   Select @Row=Count('x') from Inserted               
       Join Bag on TdBIdNo=BIdNo where Tddt<BMaxTxnDt              
   If @Row > 0              
   Begin              
    Select @ErrMsg= 'A Transaction of Later Date Exists For this Bag'              
    GoTo ErrHndlr              
   end     
   select @Row= Count('x') from inserted              
    join Bag on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo               
      and TdPrtKey=BPrtKey              
    where TdPrtKey=@wCurrPrtn and TdDt<BOpnDt              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Transaction Date Cannot Be Less Than Bag Opening Date'              
    GoTo ErrHndlr              
   end              
         


  End              

 End
   
 
 ----break1				
 --577chk          
 Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPrtn and              
  Not Exists (Select 'x' From Txn Where Txn.TIdNo=Ins.TdTIdNo and Txn.TCoCd=Ins.TdCoCd And Txn.TTc=Ins.TdTc              
   And Txn.TYy=Ins.TdYy And  Txn.TChr=Ins.TdChr And Txn.TNo=Ins.TdNo And Txn.TPrtKey=Ins.TdPrtKey)              
   If @Row> 0              
  begin              
  Select @ErrMsg= 'Voucher No. Does Not Exist'              
  GoTo ErrHndlr              
   end               
 
 --577chk          
 If @wEntTyp= 'TxnBag'              
  Begin              
   /********* Child Recs Should Have The Same Yy-Key & YyKey combination as parent ********/              
   /* Current Partition */              
   select @Row= Count('x') from inserted ins where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo<> 0              
    and exists(Select 'x' from Txnd Td1 where Td1.TdSrNo= 0              
     and (Td1.TdPrtKey=ins.TdPrtKey and Td1.TdTIdNo=ins.TdTIdNo and Td1.TdCoCd= ins.TdCoCd and Td1.TdTc= ins.TdTc              
      and Td1.TdYy= ins.TdYy and Td1.TdChr= ins.TdChr              
      and Td1.TdNo= ins.TdNo and Td1.TdSr= ins.TdSr)              
     and (Td1.TdYy<> ins.TdYy or Td1.TdKey<> ins.TdKey              
      or Td1.TdYyKey<> ins.TdYyKey))              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Yy-Key Combination Of Child Should Be Same As Parent'              
    GoTo ErrHndlr              
   end       
 end       
 --577chk TdToBloc validations
 If @wEntTyp= 'TxnBag'              
  Begin
  select @Row= Count('x') from inserted where @wTcTyp= 'BV' and TdPRtKey=@wCurrPrtn and              
    (TdFrBLoc= 'PFG' or TdToBLoc= 'PFG')              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'From/ To Bag Location Cannot Be ''PFG'' In Bag Movement'              
    GoTo ErrHndlr              
   end  
 
  IF @WTcTyp= 'Bv'              
  BEGIN              
   /* Current Partition */              
   SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn And (TdSrNo=0 AND              
    ((Ins.TdToBLoc <> ISNULL((SELECT Del.TdToBLoc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
     and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
     Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
     Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
    AND              
    NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
      LocTyp= 'P' AND LocCd= Ins.TdToBLoc AND LocValidYN = 'Y'))              
    OR              
    (TdToBLoc = '')))              
   OR              
    (TdSrNo<>0 AND TdToBLoc<> '')              
     IF @Row>0              
   BEGIN              
    SELECT @ErrMsg='Invalid Bag To Loc'              
    GOTO ErrHndlr            
   END              
              
   /* Current Partition */               
   SELECT @row= count('x') FROM inserted Ins WHERE Ins.TdPrtKey=@wCurrPrtn and Ins.TdFrBLoc = Ins.TdToBLoc              
   IF @row>0              
   BEGIN              
    SELECT @errmsg= 'To Bag Loc and From Bag Loc cannot be same'              
    GOTO ErrHndlr              
   END              
  END              
  ELSE              
  BEGIN              
   /* Current Partition */              
   SELECT @Row= COUNT('x') FROM Inserted Ins              
    WHERE TdPrtKey=@wCurrPrtn and TdToBLoc<>''              
     IF @Row>0              
   BEGIN              
    SELECT @ErrMsg='Bag To Loc should be Blank'              
      GOTO ErrHndlr              
   END              
  END          
end
--5717
 If @Mode = 'C'              
 Begin              
  /* Current Partition */                
  select @Row=Count('x') from Deleted Del where TdPrtKey=@wCurrPrtn and              
    Not Exists (Select 'x' from Inserted Ins where Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd=Del.TdCoCd              
     and Ins.TdTc=Del.TdTc and Ins.Tdyy=Del.TdYy and Ins.TdChr= Del.TdChr and Ins.TdNo=Del.TdNo              
     and Ins.TdSr=Del.TdSr and Ins.TdSrNo=Del.TdSrNo and Ins.TdPrtKey=Del.TdPrtKey              
     and Ins.TdByy=Del.TdByy and Ins.TdBChr=Del.TdBChr              
     and Ins.TdBNo=Del.TdBNo and Ins.TdKey= Del.TdKey)              
  If @Row > 0              
  Begin              
   Select @ErrMsg='Cannot Change Company Code/ Voucher No./Bag No/ Key'              
   GoTo ErrHndlr              
  End              
 End  
 --5717
 select @Row= Count(distinct(PValue)) from inserted Join param On              
    PTyp= 'TC' and PMCd= TdTc and PScd= ''              
 If @Row > 1              
 BEGIN              
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'              
  GoTo ErrHndlr              
 END 
 --5717
 select @Row= Count(distinct(PValue)) from deleted Join param on              
    PTyp='TC' and PMcd=TdTc and PSCd=''              
 If @Row > 1              
 BEGIN              
  Select @ErrMsg='Transactions of only One TcType can be Acted on'              
  GoTo ErrHndlr              
 END    
 --5717
 select @wEntTyp = PValue1 from Param where PTyp='TCTYP'              
    and PMcd=@wTcTyp and PSCd=''              
 If @wEntTyp <> 'TxnBag' and @wEntTyp <> 'TxnRm'              
 Begin              
  Select @ErrMsg='Transactions of only ''Txn'' Type permitted'              
  GoTo ErrHndlr              
 end   
 --5717
  Select @Row= Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdSr=0              
 If @Row> 0              
 Begin              
    Select @ErrMsg= 'Sr. Cannot be 0'              
    GoTo ErrHndlr              
 End
 --5717
   Select @Row= Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdKey<= 0              
 If @Row> 0              
 Begin              
  Select @ErrMsg= 'Key Cannot Be Zero'              
  GoTo ErrHndlr              
 End   
 --5717
  Select @Row=Count('x') From Inserted Ins              
  Where Ins.TdPrtKey=@wCurrPrtn and (TdToRmDc=TdFrRmDc And TdFrRmDc<>'')  Or              
   (TdFrRmDc<>'' and TdToRmDc='')              
   If @Row>0              
 Begin              
  Select @ErrMsg='Invalid To Rm Dc'              
  GoTo ErrHndlr              
 End     
 --5717
 Select @Row= count('x') from inserted where TdPrtKey=@wCurrPrtn and (TdFrRmLoc= 'FG' or TdToRmLoc= 'FG')              
  if @Row>0               
  Begin               
   Select @ErrMsg= 'Cannot Make Any Transaction For ''FG'' Location Type'              
   GoTo ErrHndlr              
  End
  --5717
  Select @Row= count('x') from inserted where TdPrtKey=@wCurrPrtn       
   and (TdFrRmLoc= 'XD' or TdToRmLoc= 'XD'       
      Or TdFrRmLocTyp In ('XD', 'MG', 'Y') or TdToRmLocTyp In ('XD', 'MG', 'Y'))      
  if @Row>0               
  Begin               
   Select @ErrMsg= 'Cannot Make Any Transaction For ''XD'',''MG'',''Y'' Location Type'              
   GoTo ErrHndlr              
  End   
  --5717
  select @Row= Count('x') from Inserted Ins where Ins.TdPrtKey=@wCurrPRtn and Ins.TdRtByQW <>'' and @wTc<>'PR'     
and @wTC<>'MV' and @wTc<>'CNV' and @wTc<>'SA' 
   If @Row> 0              
   begin              
    Select @ErrMsg= 'Rm Qw Option Should be Blank'              
    GoTo ErrHndlr              
   end
--5717
If @Mode= 'A'              
 Begin    
 select @Row= Count('x') from inserted ins where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo= 0              
   and exists(Select 'x' from Txnd Td1 where Td1.TdSrNo= 0              
    and Td1.TdCoCd= ins.TdCoCd and Td1.TdYy= ins.TdYy              
    and Td1.TdKey= ins.TdKey and              
    (Td1.TdCoCd<> ins.TdCoCd or Td1.TdTc<> ins.TdTc or              
     Td1.TdYy<> ins.TdYy or Td1.TdChr<> ins.TdChr or              
     Td1.TdNo<> ins.TdNo or Td1.TdSr<> ins.TdSr))              
    If @Row> 0              
  begin              
   Select @ErrMsg= 'Yy-Key Combination Already Exists For Another Voucher Sr.'              
   GoTo ErrHndlr              
  end 
  --5717
   If @wEntTyp= 'TxnBag'              
  Begin 
  select @Row= Count('x') from inserted ins where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo= 0              
    and exists(Select 'x' from Txnd Td1 where Td1.TdSrNo= 0              
     and Td1.TdBIdNo=ins.TdBIdNo and Td1.TdCoCd= ins.TdCoCd and Td1.TdBYy= ins.TdBYy              
     and Td1.TdBChr= ins.TdBChr  and Td1.TdBNo= ins.TdBNo              
     and (Td1.TdYy> ins.TdYy              
       Or              
      (Td1.TdYy= ins.TdYy and Td1.TdKey> ins.TdKey)))              
     If @Row> 0              
   Begin              
    Select @ErrMsg= 'A Later Txn Of The Bag Exists, Try Entering This Rec Again'              
    GoTo ErrHndlr              
   End    
   end
 end          
 --5717
 IF @wEntTyp='TxnBag'              
 Begin
 Select @Row=Count('x') From Inserted Ins              
     Where TdPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Param where PTyp='DC'              
    and PMCd = Ins.TdFrRmDc and PScd='')              
    and TdSrNo<>0)              
   OR              
    (TdSrNo=0 and TdFrRmDc<>'')              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid From Rm Ir/Rm Ir Should Not Be Entered'              
     GoTo ErrHndlr              
  END
 
 --5717
 SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and (TdSrNo<>0 AND              
   NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
      LocTyp<> 'P' and LocTyp <>'R' AND LocCd = Ins.TdToRmLoc AND LocValidYN='Y' And              
       (@wTcTyp NOT IN('Rec','Cnv','Bs', 'BM') OR LocTyp= 'Xc')))              
   OR              
    (TdSrNo=0 AND TdToRmLoc<>'')              
    IF @Row> 0              
  BEGIN              
   SELECT @ErrMsg= 'Invalid To Rm Loc/ To Rm Loc Should Be Blank?'              
   GOTO ErrHndlr              
  END    
  --5717
   SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and (TdSrNo<>0 AND              
   ((              
   (Ins.TdToRmLoc <> ISNULL((SELECT Del.TdToRmLoc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
    and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   Or              
   Ins.TdToRmDc <> ISNULL((SELECT Del.TdToRmDc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
    and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   )              
   AND TdToRmDc = 'D'              
   AND NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
      LocTyp<> 'P' and LocTyp<> 'R' AND LocCd = Ins.TdToRmLoc AND LocValidYN = 'Y' AND              
       (@wTcTyp NOT IN('Rec','Cnv','Bs','BM') OR LocTyp= 'Xc'))) OR (TdToRmLoc = '')))              
   OR              
    (TdSrNo=0 AND TdToRmLoc<>'')              
  /* **** Zubin 211 (Or Condition for 'TdToRmLoc' was missing) **** */              
    IF @Row> 0              
  BEGIN              
   SELECT @ErrMsg= 'Invalid To Rm Loc/ To Rm Loc Should Be Blank'              
   GOTO ErrHndlr              
  END        
  --5717
    Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and (@wTcTyp='MLT' and TdSrNo=0 and TdRjDc<>'C') Or              
    (@wTcTyp='REJ' and TdSrNo=0 and TdRjDc=TdBDc)              
   Or              
    ((TdSrNo<>0 or @wTctyp Not In('MLT','REJ')) and TdRjDc<>'')              
    If @Row>0             
  BEGIN              
   Select @ErrMsg='Invalid Rej Ir/Rej Ir Should Not Be Entered'              
   GoTo ErrHndlr              
  END 
  --5717
    Select @Row=Count('x') From Inserted Ins              
     Where TdPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Param where PTyp='YN'              
    and PMCd = Ins.TdAck and PScd='')              
   and  @wTcTyp='Bv')                 
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Bag Movement Ack/Ack Should Not Be Entered'              
     GoTo ErrHndlr              
  END 
  --5717
    Select @Row=Count('x') From Inserted Ins              
  Where TdPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Param where PTyp='YN'              
       and PMCd = Ins.TdPrdYN and PScd='')              
   and  @wTcTyp In('Dt') and TdSrNo=0)              
   OR              
 (@wTcTyp Not In('Dt') and TdPrdYN<> '')              
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Prd YN/Prd YN Should Not Be Entered'              
     GoTo ErrHndlr              
  END 
  --5717
    SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and              
   Ins.TdByLoc <> ISNULL((SELECT Del.TdByLoc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
     and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   AND              
   NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
     LocTyp= 'P' AND LocCd = Ins.TdByLoc AND LocValidYN = 'Y')              
   AND TdByLoc <> ''              
    IF @Row> 0              
  BEGIN              
   SELECT @ErrMsg= 'Invalid By Loc'              
     GOTO ErrHndlr              
  END  
  --5717
    SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and               
   Ins.TdByWrk <> ISNULL((SELECT Del.TdByWrk FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
    and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   AND              
     NOT EXISTS (SELECT 'x' FROM vParam WHERE vPCoCd = TdCoCd              
     AND vPTyp='WORK' AND vPMCd = Ins.TdByWrk AND vPValidYN = 'Y')              
   AND TdByWrk<>''              
      IF @Row>0              
  BEGIN              
   SELECT @ErrMsg='Invalid By Work'              
     GOTO ErrHndlr              
  END 
  --5717
    Select @Row=Count('x') From Inserted Ins              
     Where TdPrtKey=@wCurrPrtn and (TdPrdYN= 'Y' and (TdByLoc= '' or TdByWrk= '')) or              
               (IsNull( (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
     LocTyp in ('L', 'B', 'M')  and LocCd = TdToRmLoc), '')<> '' and (TdByLoc= '' or TdByWrk= '')              
    )              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='By Loc/ Wrk Should Be Entered'              
     GoTo ErrHndlr              
  END   
  --5717
    Select @Row= Count('x') from inserted Ins       
  Join Txn On TPrtKey=Ins.TdPRtKey and TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where TdPrtKey=@wCurrPrtn and              
   (@wTcTyp not in ('BS', 'BM') or TdSrNo<> 0 or (((TMultIIr<>'I' and TdBDc<> 'D') or (TMultIIr='I' and TdBDc<> 'C')) and @wTcTyp<> 'BM') ) and               
   (TdRefYy<> '' or TdRefKey<> 0)                 
      If @Row > 0               
       Begin              
   Select @ErrMsg='Ref Yy && Ref Key should be blank'              
     GoTo ErrHndlr              
    end  
	--5717
	 Select @Row=Count('x') from inserted Ins where TdPrtKey=@wCurrPrtn and              
   (Not Exists (Select 'x' from Txnd  where Txnd.TdPrtKey=Ins.TdPrtKey and Txnd.TdBIdNo=Ins.TdBIdNo and Txnd.TdCoCd=Ins.TdCoCd and              
    Txnd.TdBYy=Ins.TdBYy and Txnd.TdBChr=Ins.TdBChr and Txnd.TdBNo=Ins.TdBNo              
    and Txnd.TdYy=Ins.TdPrevYy and Txnd.TdKey=Ins.TdPrevKey)              
   and TdSrno=0)              
   and              
   (TdPrevYy<>'' or TdPrevKey<>0)              
    If @Row > 0              
     Begin              
   Select @ErrMsg='Invalid Prev YY/Key'              
   GoTo ErrHndlr              
  end
  --5717
    if @mode in ('A','C')              
  begin                 
   Select @Row=Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and              
    Exists (Select 'x' From Bag Where BPrtKey=Ins.TdPrtKey and BIdNo=Ins.TdBIdNo and BCoCd= Ins.TdCoCd and              
     BYy= Ins.TdByy and BChr= Ins.TdBChr and BNo= Ins.TdBNo and Cast(BQty as Decimal(8,2))< 0)              
     If @Row>0              
   begin              
    Select @ErrMsg='Cannot accept Txn, Bag Qty going -ve'              
    GoTo ErrHndlr              
   end              
  end              
  --5717            
  if @Mode='D'              
  Begin                 
   Select @Row=Count('x') From Deleted Del Where Del.TdPRtKey=@wCurrPrtn and               
    Exists (Select 'x' From Bag Where BPrtKey=Del.TdPrtKey and BIdNo=Del.TdBIdNo and BCoCd= Del.TdCoCd              
         and BYy= Del.TdByy and BChr= Del.TdBChr              
         and BNo = Del.TdBNo and Cast(BQty as Decimal(8,2))< 0)              
     If @Row>0              
     begin              
    Select @ErrMsg='Cannot accept Txn, Bag Qty going -ve'              
    GoTo ErrHndlr              
   end              
  End
  --5717
    Select @Row=Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPrtn and Ins.TdRmCd<>'' and              
   IsNull((Select sum(Td.TdRmQty* (              
        Case when Td.TdFrRmLoc= 'W'  and Td.TdFrRmDc= 'D' then 1              
     when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
     when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
     when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
            Else 0 end))              
     from Txnd Td where Td.TdPrtKey=Ins.TdPrtKey and Td.TdBIdNo=Ins.TdBIdNo               
     and Td.TdCoCd= Ins.TdCoCd and Td.TdBYy= Ins.TdBYy              
     and Td.TdBChr= Ins.TdBChr and Td.TdBNo= Ins.TdBNo              
     and Td.TdRmCd= Ins.TdRmCd and Td.TdLotNo= Ins.TdLotNo              
     and Str(Td.TdRmSz, 16, 4)= Str(Ins.TdRmSz, 16, 4)              
     and Str(Td.TdRmSz2, 16, 4)= Str(Ins.TdRmSz2, 16, 4)              
     and Str(Td.TdRmSz3, 16, 4)= Str(Ins.TdRmSz3, 16, 4)              
     and Str(Td.TdRmStkRt, 16, 2)= Str(Ins.TdRmStkRt, 16, 2)), 0)< 0              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Rm Qty Is Less Than Zero In Bag'              
   GoTo ErrHndlr              
  END    
  --5717
  If @Mode in ('C')              
  Begin              
   Select @Row= Count('x') From Deleted Del Where Del.TdPrtKey=@wCurrPrtn and Del.TdRmCd<> '' and               
    exists(Select 'x' from Inserted Ins where               
  Ins.TdPrtKey=Del.TdPRtKey and Ins.TdBIdNo=Del.TdBIdNo              
     and Ins.TdCoCd= Del.TdCoCd and Ins.TdBYy= Del.TdBYy              
     and Ins.TdBChr= Del.TdBChr and Ins.TdBNo= Del.TdBNo              
     and (Ins.TdRmCd<> Del.TdRmCd or Ins.TdLotNo<> Del.TdLotNo              
     or Str(Ins.TdRmSz, 16, 3)<> Str(Del.TdRmSz, 16, 3)              
     or Str(Ins.TdRmSz2, 16, 3)<> Str(Del.TdRmSz2, 16, 3)              
     or Str(Ins.TdRmSz3, 16, 3)<> Str(Del.TdRmSz3, 16, 3)              
     or Str(Ins.TdRmStkRt, 16, 2)<> Str(Del.TdRmStkRt, 16, 2))) and              
     Round(IsNull((Select sum(Td. TdRmQty* (              
     Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
     Else 0 end))              
    from Txnd Td where Td.TdPrtKey=Del.TdPRtKey and Td.TdBIdNo=Del.TdBIdNo              
    and Td.TdCoCd= Del.TdCoCd and Td.TdBYy= Del.TdBYy              
    and Td.TdBChr= Del.TdBChr and Td.TdBNo= Del.TdBNo              
    and Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
    and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
    and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
	and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
    and Str(Td.TdRmStkRt, 16, 2)= Str(Del.TdRmStkRt, 16, 2)), 0), 3)< 0               
  If @Row>0              
  Begin              
   Select @ErrMsg='Rm Qty Is Less Than Zero In Bag'              
   GoTo ErrHndlr              
  End            
  End

  --5717
  if @Mode='D'              
  Begin              
   Select @Row=Count('x') From Deleted Del Where Del.TdPrtKey=@wCurrPrtn and Del.TdRmCd<>'' and              
    IsNull((Select sum(Td.TdRmQty* (              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end))            
      from Txnd Td where Td.TdPrtKey=Del.TdPrtKey and Td.TdBIdNo=Del.TdBIdNo              
      and Td.TdCoCd= Del.TdCoCd and Td.TdBYy= Del.TdBYy              
      and Td.TdBChr= Del.TdBChr and Td.TdBNo= Del.TdBNo              
      and Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
      and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
	  and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
	  and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
      and Str(Td.TdRmStkRt, 16, 2)= Str(Del.TdRmStkRt, 16, 2)), 0)< 0              
     If @Row>0              
   begin              
    Select @ErrMsg='Rm Qty Is Less Than Zero In Bag'              
    GoTo ErrHndlr              
   end           
   End
   --5717
    Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPrtn and Ins.TdRmCd<> '' and              
   Round(IsNull((Select sum(Td.TdRmWt* (              
        Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
     when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
     when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
     when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
            Else 0 end))              
     from Txnd Td where Td.TdPrtKey=Ins.TdPrtKey and Td.TdBIdNo=Ins.TdBIdNo              
     and Td.TdCoCd= Ins.TdCoCd and Td.TdBYy= Ins.TdBYy              
     and Td.TdBChr= Ins.TdBChr and Td.TdBNo= Ins.TdBNo              
     and Td.TdRmCd= Ins.TdRmCd and Td.TdLotNo= Ins.TdLotNo              
     and Str(Td.TdRmSz, 16, 4)= Str(Ins.TdRmSz, 16, 4)              
     and Str(Td.TdRmSz2, 16, 4)= Str(Ins.TdRmSz2, 16, 4)              
	 and Str(Td.TdRmSz3, 16, 4)= Str(Ins.TdRmSz3, 16, 4)              
     and Str(Td.TdRmStkRt, 16, 2)= Str(Ins.TdRmStkRt, 16, 2)), 0), 3)< 0              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Rm Wt Is Less Than Zero In Bag'              
   GoTo ErrHndlr              
  END    
  --5717
    If @Mode in ('C')              
  Begin              
   Select @Row= Count('x') From Deleted Del Where Del.TdPrtKey=@wCurrPrtn and Del.TdRmCd<> '' and               
    exists(Select 'x' from Inserted Ins where            
     Ins.TdPrtKey=Del.TdPRtKey and Ins.TdBIdNo=Del.TdBIdNo              
     and Ins.TdCoCd= Del.TdCoCd and Ins.TdBYy= Del.TdBYy              
     and Ins.TdBChr= Del.TdBChr and Ins.TdBNo= Del.TdBNo              
     and (Ins.TdRmCd<> Del.TdRmCd or Ins.TdLotNo<> Del.TdLotNo              
     or Str(Ins.TdRmSz, 16, 3)<> Str(Del.TdRmSz, 16, 3)              
     or Str(Ins.TdRmSz2, 16, 3)<> Str(Del.TdRmSz2, 16, 3)              
     or Str(Ins.TdRmSz3, 16, 3)<> Str(Del.TdRmSz3, 16, 3)              
     or Str(Ins.TdRmStkRt, 16, 2)<> Str(Del.TdRmStkRt, 16, 2))) and              
     Round(IsNull((Select sum(Td.TdRmWt* (              
     Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
     Else 0 end))              
    from Txnd Td where Td.TdPrtKey=Del.TdPRtKey and Td.TdBIdNo=Del.TdBIdNo              
    and Td.TdCoCd= Del.TdCoCd and Td.TdBYy= Del.TdBYy              
    and Td.TdBChr= Del.TdBChr and Td.TdBNo= Del.TdBNo              
    and Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
    and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
    and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
    and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
    and Str(Td.TdRmStkRt, 16, 2)= Str(Del.TdRmStkRt, 16, 2)), 0), 3)< 0               
  If @Row>0              
  Begin              
   Select @ErrMsg='Rm Wt Is Less Than Zero In Bag'              
   GoTo ErrHndlr              
  End            
  End
  --5717
    if @Mode='D'              
  Begin              
   Select @Row= Count('x') From Deleted Del Where Del.TdPrtKey=@wCurrPrtn and Del.TdRmCd<> '' and              
    Round(IsNull((Select sum(Td.TdRmWt* (              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end))              
      from Txnd Td where Td.TdPrtKey=Del.TdPRtKey and Td.TdBIdNo=Del.TdBIdNo              
      and Td.TdCoCd= Del.TdCoCd and Td.TdBYy= Del.TdBYy              
      and Td.TdBChr= Del.TdBChr and Td.TdBNo= Del.TdBNo              
      and Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
      and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
      and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
      and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
      and Str(Td.TdRmStkRt, 16, 2)= Str(Del.TdRmStkRt, 16, 2)), 0), 3)< 0              
     If @Row>0              
   begin              
    Select @ErrMsg='Rm Wt Is Less Than Zero In Bag'              
    GoTo ErrHndlr              
   end        
   End
   --5717
     Select @Row=Count('x') From Inserted Ins              
   Where TdPRtKey=@wCurrPrtn               
   and Not Exists (Select 'x' From Bag Where BPrtKey=Ins.TdPrtKey and BIdNo=Ins.TdBIdNo              
       and BCoCd= Ins.TdCoCd              
       and BYy= Ins.TdByy and BChr=Ins.TdBChr              
       and BNo = Ins.TdBNo and BRjQty>= 0)              
    If @Row>0              
  BEGIN              
   Select @ErrMsg= 'Cannot accept Bag Rej Qty going -ve'              
   GoTo ErrHndlr              
  END     
 end 
  --5717 
  If @wEntTyp= 'TxnBag'              
  BEGIN    
    select @Row= Count('x') from inserted Ins       
  Join Txn On TPrtKey=Ins.TdPRtKey and TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where @wTcTyp= 'BS' and TdPRtKEy=@wCurrPrtn and TMultiIr<>'I'  and            
   (TdRefYy<> '' or TdRefKey<> 0) and               
   Not Exists (Select 'x' from Txnd Td where Td.TdPrtKey=Ins.TdPrtKey and Td.TdTIdNo=Ins.TdTIdNo               
   and Td.TdCoCd= Ins.TdCoCd and Td.TdTc= Ins.TdTc and Td.Tdyy= Ins.TdYy               
   and Td.TdChr= Ins.TdChr and Td.TdNo= Ins.TdNo               
   and Td.TdYy= Ins.TdRefYy and Td.TdKey= Ins.TdRefKey)      
  /*** (Jen 2.13) ***/              
  If @Row > 0              
  Begin              
   Select @ErrMsg= 'RefYy, RefKey is not in the Voucher'              
   GoTo ErrHndlr              
  end   
  --5717
    select @Row=Count('x') from inserted Ins where Ins.TdPrtKey=@wCurrPrtn and Ins.TdSrNo<>0 and              
   Not Exists (Select 'x' from Txnd Td               
   where Td.TdPRtkey=Ins.TdPrtKey and Td.TdTIdNo=Ins.TdTIdNo               
   and Td.TdCoCd=Ins.TdCoCd and Td.TdTc=Ins.TdTc and Td.Tdyy=Ins.Tdyy              
   and Td.TdChr= Ins.TdChr and Td.TdNo=Ins.TdNo              
   and Td.TdSr=Ins.TdSr and Td.TdSrNo=0              
   and Td.TdBIdNo=Ins.TdBIdNo and Td.TdBYy= Ins.TdBYy and Td.TdBChr= Ins.TdBChr              
   and Td.TdBNo= Ins.TdBNo)              
  If @Row > 0              
  Begin              
   Select @ErrMsg='TdSrNo=0 record does not exist'              
   GoTo ErrHndlr              
  end 
  --5717
    If @wTcTyp= 'BS'              
  Begin   
     Select @Row= Count('x') From Inserted Ins              
    Join Param P1 On P1.PTyp= 'BCHR' and P1.PMCd= ins.TdBChr and P1.PSCd= ''              
    where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo= 0 and          
     Exists (Select 'x' from Txnd Td              
      Join Param P2 On P2.PTyp= 'BCHR' and P2.PMCd= Td.TdBChr and P2.PSCd= ''              
      where Td.TdPRtKey=Ins.TdPRtKey and Td.TdTIdNo=Ins.TdTIdNo and Td.TdCoCd= Ins.TdCoCd and              
      Td.TdTc= Ins.TdTc and Td.TdYy= Ins.TdYy and Td.TdChr= Ins.TdChr and              
      Td.TdNo= Ins.TdNo and Td.TdSrNo= 0 and P2.PValue<> P1.PValue )              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Bag Split Can Take Place Between Same Bag Type'              
    GoTo ErrHndlr              
   end       
   --5717
      select @Row= Count('x') from inserted Ins      
  Join Txn On TPrtKey=Ins.TdPRtKey and TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and ((TMultiIr<>'I' and TdBDc= 'D') /*or (TMultiIr='I' and TdBDc= 'C')*/)      
    and (TdRefYy= '' or TdRefKey<= 0)              
       If @Row > 0              
        begin              
    Select @ErrMsg= 'RefYy / RefKey cannot be Blank'              
      GoTo ErrHndlr              
     end   
	   
	end
  END
              
              
 If @Mode ='D' and @wEntTyp='TxnBag'              
  Begin              
   Select @Row=Count('x') from Deleted Del where TdSrNo=0 and              
      Exists (Select 'x' from Txnd Td Where Td.TdTIdNo=Del.TdTIdNo               
    And Td.TdCoCd= Del.TdCoCd And Td.TdTc= Del.TdTc and Td.Tdyy= Del.Tdyy              
    And Td.TdChr= Del.TdChr and Td.TdNo= Del.TdNo              
    And Td.TdSr= Del.TdSr and Td.TdSrNo<> 0)              
   If @Row > 0              
   Begin              
    Select @ErrMsg='Cannot Delete SrNo=0 Record, As Childs Exist'              
    GoTo ErrHndlr              
   End              
              
   Select @Row=Count('x') from Deleted Del where  TdSrNo=0 and              
      Exists (Select 'x' from TxndPrd where TpTdIdNo=TdIdNo and TpCoCd= TdCoCd and TpTc= TdTc              
     and TpYy= TdYy and TpChr= TdChr and TpNo= TdNo and TpSr= TdSr)              
   If @Row > 0              
   Begin              
    Select @ErrMsg='Cannot Delete SrNo=0 Record, As Prd Points Record Exists'              
    GoTo ErrHndlr              
   End           
  End              
  --5717
  if @whOLORYN= 'Y'
 begin
  select @Row= Count('x') from inserted ins where TdRmCtg in ('D','C','X') and TdLotNo =''              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Lot No. should Be Entered For Diamonds,Color Stones,Accessories'              
   GoTo ErrHndlr              
  End
  --5717
  select @Row= Count('x') from inserted ins where TdRmCtg in ('D','C','X') and @wEntTyp<> 'TxnBag' and TdLotNo <>'' and TdPurRt=0             
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Rate cannot be zero, Rate should Be Entered For Diamonds,Color Stones,Accessories'              
   GoTo ErrHndlr              
  End
  --5717
  Select @Row = Count('x') from inserted ins where TdRmCtg not in ('D','C','X') and TdLotNo <> ''              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Lot No. should not be entered for other than Diamonds,Color Stones,Accessories'              
   GoTo ErrHndlr              
  End
  --5717
  Select @Row=Count('x') from Deleted Del where TdLotMstRecYN='Y' and @mode='D' and                
      Exists (Select 'x' from Txnd where TdLotNo = Del.TdLotNo and TdIdNo <> Del.TdIdNo)              
   If @Row > 0              
   Begin              
    Select @ErrMsg='Cannot Delete Lot Master record as child transactions exist for this Lot No'              
    GoTo ErrHndlr              
   End  
   --5717
   If @Mode = 'C'              
   Begin                 
    select @Row=Count('x') from Deleted Del where TdLotMstRecYN='Y' and exists (Select 'x' from Txnd where TdLotNo = Del.TdLotNo and TdIdNo <> Del.TdIdNo)             
        If @Row > 0              
   Begin              
    Select @ErrMsg='Cannot Change RmCd/LotNo/RtByQW/Rt for a Lot Master record as child transactions exist for this Lot No'              
    GoTo ErrHndlr              
   End              
  End    
End
 --5717
  If @wEntTyp in ('TxnRm', 'TxnBag')              
 Begin                
  select @Row= Count('x') from inserted ins  where @whOLORYN ='N' and TdPRtKey=@wCurrPrtn and TdLotNo <> '' and              
   Not Exists (select 'x' from RmMst where RmZ= 'N' and RmPrtKey=TdPrtKey and RmCd= TdRmCd and RmCtg IN ('D', 'C', 'X','G','P','S','L','M'))              
    If @Row> 0              
  begin              
   Select @ErrMsg= 'Lot No. should Be Entered For Diamonds ,Color Stones,Accessories, Gold, Palladium, Silver, Platinum and Metal Only'              
   GoTo ErrHndlr              
  end  
  --5717
  Select @Row= Count('x') From Inserted Ins Where @wholoryn='N' and TdPrtKey=@wCurrPRtn and TdLotNo<> '' and              
   Not Exists (Select 'x' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and HLotNoYN= 'Y')              
    If @Row> 0              
  begin              
   Select @ErrMsg= 'Lot No. should Be Blank As ''Implement Lot No. '' Option Is ''N'' In Head Master (Tab 5)'              
     GoTo ErrHndlr              
  end              
 End      
    

  If @wTcTyp In('DT', 'BS')              
  Begin              
	if @Mode='A'
	begin
		Select @Row= Count('x') from Inserted Ins              
			Join RmMst On RmCd= Ins.TdRmCd and RmPrtKey=Ins.TdPrtKey              
			Join Param ON PTyp= 'BCHR' and PMCd= Ins.TdBChr and PSCd= '' and PValue in ('F')               
			where Ins.TdPrtKey=@wCurrPRtn and Ins.TdSrNo> 0 and              
				NOT (              
				 (',' + PVALUE1 + ',') LIKE (CASE WHEN PVALUE1='' THEN '%' ELSE '%,'+RmCtg+',%' END) and              
				 (',' + PVALUE2 + ',') LIKE (CASE WHEN PVALUE2='' THEN '%' ELSE '%,'+RmSCtg+',%' END) and
				 ((Pdesc225='' and PValue3='') or (exists (Select 'a' from Bag 
																		join OrdRm on Bag.BOdIdNo = OrdRm.OrOdIdNo 
																		where Bag.BCoCd=Ins.TdCoCd and Bag.BYy = Ins.TdByy and Bag.BChr = Ins.TdBChr and Bag.BNo = Ins.TdBNo and TdRmCd =OrRmCd and
																		(PDesc225='' or PDesc225=OrSubShp)
																		and (PValue3='' or (PValue3='WS' and OrWsQty>0) or (PValue3='GS' and OrHsQty>0))
																		 					))))
				
	   If @Row> 0              
		begin              
			Select @ErrMsg= 'Cannot Issue Raw Material Of This Ctg/Sub-Category To This Flute Bag'              
			GoTo ErrHndlr              
		end              
	end
 End              
 
   If @wTcTyp In('DT')              
  Begin              
   --5717
   select @Row= Count('x') from Inserted Ins              
    where Ins.TdPrtKey=@wCurrPrtn and (Ins.TdFBRmQty<> '' or Ins.TdFBRmWt<> '' ) and              
     Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= Ins.TdBChr and              
      PSCd= '' and PValue in ('C', 'P'))              
   If @Row> 0              
   begin              
    Select @ErrMsg= 'Cannot Enter RM Quantity/Weight Reference For Primary/ Component Bags'              
    GoTo ErrHndlr              
   end
   /* Zubin 212 */              
   /* Current Partition */              
	if @wTChr = 'Y'
	begin
     select @row= count('x') from inserted where TdPrtKey=@wCurrPRtn and TdSrNo <> 0 and              
     TdFrRmLoc<> (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'XR') AND              
     TdToRmLoc<> (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'XR')              
     if @row> 0              
		begin              
			select @errmsg= 'Either RmLocFr Or RmLocTo Should Belong To ''XR'' Loc Type'              
			goto ErrHndlr              
		end              


     select @row= count('x') from inserted where TdPrtKey=@wCurrPRtn and TdSrNo <> 0 and              
     TdFrRmLoc= (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'W') Or              
     TdToRmLoc= (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'W')              
     if @row> 0              
		begin              
			select @errmsg= 'RmLocFr Or RmLocTo cannot Belong To ''W'' Loc Type'              
			goto ErrHndlr              
		end              
	END
	else
	begin
	     select @row= count('x') from inserted where TdPrtKey=@wCurrPRtn and TdSrNo <> 0 and              
			TdFrRmLoc<> (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'W') AND              
			TdToRmLoc<> (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'W')              
			if @row> 0              
			begin              
				select @errmsg= 'Either RmLocFr Or RmLocTo Should Belong To ''W'' Loc Type'              
				goto ErrHndlr              
			end              
	
	 select @row= count('x') from inserted where TdPrtKey=@wCurrPRtn and TdSrNo <> 0 and              
     TdFrRmLoc= (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'XR') Or              
     TdToRmLoc= (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'XR')              
     if @row> 0              
		begin              
			select @errmsg= 'RmLocFr Or RmLocTo cannot Belong To ''XR'' Loc Type'              
			goto ErrHndlr              
		end              
	ENd
   /* Zubin 212 */              
  END
 --5717
   If @wTcTyp Not In('DT')              
  Begin              
   /* Current Partition */               
   select @Row= Count('x') from Inserted where TdPrtKey=@wCurrPrtn and TdFBRmQty<> '' or TdFBRmWt<> ''              
   If @Row> 0              
   begin              
    Select @ErrMsg= 'Cannot Enter RM Quantity/Weight Reference For This TcTyp'              
    GoTo ErrHndlr              
   end              
  End   
  --5717
  if @mode = 'C'              
  Begin
  If @wTcTyp= 'DT'              
   begin              
    /* Current Partition */              
    select @Row= Count('x') from Inserted Ins where TdPRtKey=@wCurrPrtn and   
     Exists (Select 'x' from deleted del where ins.TdPrtKey=del.TdPrtKey and ins.TdIdNo=del.TdIdNo               
     and ins.TdCoCd= del.TdCoCd and ins.TdTc= del.TdTc and ins.TdYy= del.TdYy              
     and ins.TdChr= del.TdChr and ins.TdNo= del.TdNo              
     and ins.TdSr = del.TdSr and ins.TdSrNo = del.TdSrNo and              
      (ins.TdByWrk<> del.TdByWrk or ins.TdByLoc<> del.TdByLoc))              
     and              
     Exists (Select 'x' from TxndPrd where TpPrtKey=ins.TdPrtKey               
      and TpTdIdNo=ins.TdIdNo and TpCoCd= ins.TdCoCd              
      and TpTc= ins.TdTc and TpYy= ins.TdYy and TpChr= ins.TdChr              
      and TpNo= ins.TdNo and TpSr= ins.TdSr)              
    If @Row > 0              
    begin              
     Select @ErrMsg= 'Cannot Change Prd as Production Points Exist'              
     GoTo ErrHndlr              
    end              
   end         
   select @Row= Count('x') from inserted ins where ins.TdPrtKey=@wCurrPrtn and ins.TdTc = 'BV' and              
   not exists(Select 'X' from deleted del where del.TdPrtKey=ins.TdPrtKey and del.TdIdNo=ins.TdIdNo               
   and del.Tdcocd= ins.TdCoCd and del.TdTc= ins.TdTc and del.TdYy= ins.TdYy and del.TdChr= ins.TdChr              
   and del.TdNo= ins.TdNo and del.TdSr= ins.TdSr and del.TdBYy= ins.TdBYy              
   and del.TdBChr= ins.TdBChr and del.TdBNo= ins.TdBNo and del.TdBQty= ins.tdBQty              
   and del.TdBDc= ins.TdBDc and del.TdBGrWt= ins.TdBGrWt and del.TdFrBLoc= ins.TdFrBLoc              
   and del.TdToBLoc= ins.TdToBLoc and del.TdFrRmLoc= ins.TdFrRmLoc and del.TdFrRmDc= ins.TdFrRmDc              
   and del.TdRmCd= ins.TdRmCd and del.TdLotNo= ins.TdLotNo and del.TdRmSz= ins.TdRmSz              and
   del.TdRmSz2= ins.TdRmSz2  and del.TdRmSz3= ins.TdRmSz3                          
   and del.TdRmStkRt= ins.TdRmStkRt              
   and del.TdRmQty= ins.TdRmQty and del.TdRmWt= ins.TdRmWt and del.TdToRmLoc= ins.TdToRmLoc              
   and del.TdToRmDc= ins.TdToRmDc and del.TdWrk= ins.TdWrk and del.TdPrdYN= ins.TdPrdYN              
   and del.TdByLoc= ins.TdByLoc and del.TdByWrk= ins.TdByWrk and del.TdDustWt= ins.TdDustWt and              
   del.TdKey= ins.TdKey and del.TdRefYy= ins.TdRefYy and del.TdRefKey= ins.TdRefKey and              
   del.TdRjCd= ins.TdRjCd and del.TdRjQty= ins.TdRjQty and del.TdRjDc= ins.TdRjDc and              
   del.TdPtQty= ins.TdPtQty and del.TdRjPtQty= ins.TdRjPtQty and del.TdDesc= ins.TdDesc              
   and del.TdPurRt= ins.TdPurRt and del.TdPurAmt= ins.TdPurAmt and del.TdPrevYy= ins.TdPrevYy              
   and del.TdPrevKey= ins.TdPrevKey and del.TdYyKey= ins.TdYyKey and del.TdPcPerCt= ins.TdPcPerCt )              
   If @Row > 0              
   Begin              
    Select @ErrMsg= 'Cannot Update any as Bag is Acknowledged'              
    GoTo ErrHndlr              
   end              
  End 
  
  
             
 If @wEntTyp= 'TxnBag'              
 Begin              
          
              
  /**** Cannot Update or Delete if Fg Txn later to the Txn Entry of the Bag exists *******/              
  /* Current Partition */               
  if @mode='D'
  begin 
   Select @Row=Count('x') from deleted where 
   Exists (Select 'x' from Fgd where FdBIdNo=TdBIdNo and FdCoCd= TdCoCd and FdBYy= TdBYy              and FdBChr= TdBChr and FdBNo= TdBNo and FdBYyKey>= TdYyKey)              
  If @Row > 0              
  begin              
   Select @ErrMsg= 'Cannot Update Or Delete, A Later Fg Txn Exists For The Bag'              
   GoTo ErrHndlr              
  end
  end
  
  if @mode in ('A','C')
  begin 
	Select @Row=Count('x') from inserted where TdSrNo<> 0 and              
    Exists (Select 'x' from Fgd where FdBIdNo=TdBIdNo and FdCoCd= TdCoCd and FdBYy= TdBYy              
   and FdBChr= TdBChr and FdBNo= TdBNo and FdBYyKey>= TdYyKey)              
  If @Row > 0              
  begin              
   Select @ErrMsg= 'Cannot Add Or Update, A Later Fg Txn Exists For The Bag'              
   GoTo ErrHndlr              
  end
  
  Select @Row=Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and              
   Not Exists (Select 'x' From Loc where LocCoCd= TdCoCd and              
   LocTyp= 'P' and LocCd = Ins.TdFrBLoc)              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Bag Fr Loc, contact 9820490905 right away'              
     GoTo ErrHndlr              
  END              
  Select @Row=Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and              
   (Ins.TdFrBLoc <> ISNULL((SELECT Del.TdFrBLoc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
    and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   AND NOT EXISTS (Select 'x' From Loc where LocCoCd= TdCoCd and              
    LocTyp= 'P' and LocCd = Ins.TdFrBLoc AND LocValidYN = 'Y')              
   OR Ins.TdFrBLoc = '')              
   AND @wTcTyp = 'REJ' And TdBDc = 'D'              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Bag Fr Loc'              
     GoTo ErrHndlr              
  END              
  --5717
  Select @Row=Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and              
   Not Exists (Select 'x' From Loc where LocCoCd= TdCoCd and              
   LocTyp= 'P' and LocCd = Ins.TdFrBLoc)              
    If @Row>0              
	BEGIN              
		Select @ErrMsg='Invalid Bag Fr Loc'              
		GoTo ErrHndlr              
	END    
    end          
              
  if @mode = 'A'              
  Begin              
   select @Row=Count('x') from inserted where @wTcTyp= 'MLT' and
    exists (Select Cast(BQty As Decimal(8,1)) from Bag where TdBIdNo=BIdNo and (cast(Bqty as decimal(8,1))> 0 or BPtQty<>0))
   If @Row > 0              
   begin              
    Select @ErrMsg='Cannot Melt Bag As Bag Qty/Part Qty <> 0'              
    GoTo ErrHndlr              
   end              
              
	
   select @Row=Count('x') from inserted ins join Bag on BIdNo=TdBIdNo               
    where @wTcTyp not in('MLT', 'REJ') and (Round(BRjQty, 1)> 0 or BRjPtQty<> 0)              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Make Any Transaction For A Rejected Bag'              
    GoTo ErrHndlr              
   end              
   --5717
   select @Row=Count('x') from inserted ins where TdPrtKey=@wCurrPrtn and              
    exists (Select 'x' from Txnd Td where Td.TdPrtKey=ins.TdPrtKey and               
     Td.TdCoCd= ins.TdCoCd and Td.TdYy= ins.TdPrevYy and Td.TdKey= ins.TdPrevKey and              
     (Select PValue from Param              
      where PTyp= 'TC' and PMCd= Td.TdTc)= 'MLT')              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Make Any Transaction For A Melted Bag'              
    GoTo ErrHndlr              
   end              
 End              
   --5717
   if @wTcTyp in ('BV', 'REJ')              
  begin             /* Current Partition */              
   Select @Row= Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdLotNo<> ''              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Lot No. Should Be Blank'              
      GoTo ErrHndlr              
   end              
  end
  --5717
  select @Row=Count('x') from inserted ins where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and TdLotNo<> ''              
  If @Row > 0              
  begin              
   Select @ErrMsg= 'Lot No Should Be Blank For TdSrNo= 0'              
   GoTo ErrHndlr              
  end 
             
 END      
--5717
 if @mode in ('A','C') and @wEntTyp= 'TxnRm'              
 Begin              
   If @wTcTyp= 'PR'               
   Begin                
  /***** Either All Best Use Order Fields are filled up or all are blank *****/        
  Select @Row= Count('x') From Inserted Ins       
  Where Ins.TdPrtKey=@wCurrPrtn and (TdBstOdTc<> '' or TdBstOdYy<> '' or TdBstOdChr<> '' or TdBstOdNo<> 0 )      
   and (TdBstOdTc= '' or TdBstOdYy= '' or TdBstOdChr= '' or TdBstOdNo= 0 )       
  If @Row> 0      
  begin      
     Select @ErrMsg= 'Either All Order Fields Should Be Filled Or All Should Be Blank'      
     GoTo ErrHndlr       
  end     
  --5717
  Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPRtn and TdBstOdNo<> 0       
   and not exists(Select 'x' From OrdMst where OmPrtKey=TdPrtkey and OmCoCd= TdCoCd and OmTc= TdBstOdTc and OmYy= TdBstOdYy       
    and OmChr= TdBstOdChr and OmNo= TdBstOdNo)       
  If @Row> 0       
  begin      
     Select @ErrMsg= 'Order Number Not Defined'      
     GoTo ErrHndlr      
  end 
  --5717
  Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPRtn and TdBstOdSr <> 0       
   and not exists(Select 'x' From OrdDsg where OdPrtKey=TdPrtkey and OdCoCd= TdCoCd and OdTc= TdBstOdTc and OdYy= TdBstOdYy       
    and OdChr= TdBstOdChr and OdNo= TdBstOdNo and OdSr =TdBstOdSr )       
  If @Row> 0       
  begin      
     Select @ErrMsg= 'Order Serial Number Not Defined'      
     GoTo ErrHndlr      
  end  
  --5717
  Select @Row= Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdBstDmCd<> ''       
   and not exists(Select 'x' From DsgMst where DmPrtKey=TdPrtKey and DmTcTyp= 'DM' and DmCd= TdBstDmCd and DmSz= '')       
  If @Row> 0       
  begin      
     Select @ErrMsg= 'Design Code Not Defined'      
     GoTo ErrHndlr      
  end    
  End      
 End 


if @wTcTyp ='BM' and @wEntTyp ='TxnBag'
begin
	select @Row= Count('x') from inserted Ins where              
    Not Exists (Select 'x' from Txnd where Txnd.TdTIdNo=Ins.TdTIdNo              
      and Txnd.TdCoCd= Ins.TdCoCd             
         and Txnd.TdSrNo= 0 and Txnd.TdBDc= 'C' and Txnd.TdRefKey= Ins.TdRefKey and Txnd.TdRefYy =0)               
    and TdBDc= 'D' and TdSrNo= 0              
     If @Row> 0              
        Begin            
    Select @ErrMsg= 'No Parent Bag For The Child Bag'              
      GoTo ErrHndlr              
     end        
	    select @Row= Count('x') from inserted Ins               
    Join Txnd Td On Td.TdPrtKey=Ins.TdPrtKey and Td.TdTIdNo=Ins.TdTIdNo              
     and Td.TdCoCd= Ins.TdCoCd and Td.TdTc= Ins.TdTc and Td.TdYy= Ins.TdYy and Td.TdChr= Ins.TdChr and               
     Td.TdNo= Ins.TdNo and Td.TdSrNo= 0 and Td.TdRefKey= Ins.TdRefKey               
    where Ins.TdSrNo= 0 and               
     not exists (select 'x' from Bag B1               
      join Bag B2 on B1.BPrtKey=B2.BPrtKey and B1.BOdIdNo=B2.BOdIdNo              
       and B1.BCoCd= B2.BCoCd and B1.BOdTc= B2.BOdTc               
       and B1.BOdYy= B2.BOdYy and B1.BOdChr= B2.BOdChr               
       and B1.BOdNo= B2.BOdNo and B1.BOdSr= B2.BOdSr               
      where B1.BPrtKey=Td.TdPrtKey and B1.BIdNo=Td.TdBIdNo and B1.BCoCd= Td.TdCoCd               
       and B1.BYy= Td.TdBYy and B1.BChr= Td.TdBChr and B1.BNo= Td.TdBNo               
       and B2.BCoCd= Ins.TdCoCd and B2.BYy= Ins.TdBYy               
       and B2.BChr= Ins.TdBChr and B2.BNo= Ins.TdBNo)               
        If @Row> 0               
        Begin              
    Select @ErrMsg= 'Order Design Serial Not Same As That Of Previous Bag'               
      GoTo ErrHndlr               
        end              

	    --577 chk -it is commented in 577 trigger, so now in 5715 also commented
		--select @Row= Count('x') from inserted Ins              
  --  Join Txnd Td On Td.TdCoCd= Ins.TdCoCd and              
  --   Td.TdTc= Ins.TdTc and Td.TdYy= Ins.TdYy and Td.TdChr= Ins.TdChr and              
  --   Td.TdNo= Ins.TdNo and Td.TdSrNo= 0              
  --  where Ins.TdSrNo= 0 and              
  --   (select BOdDmCd from Bag where BCoCd= Td.TdCoCd and              
  --    BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo)<>              
  --   (select BOdDmCd from Bag where BCoCd= Ins.TdCoCd and              
  --    Byy= Ins.TdBYy and BChr= Ins.TdBChr and BNo= Ins.TdBNo)              
  --      If @Row> 0              
  --      Begin              
  --  Select @ErrMsg= 'Design Code Not Same As That Of Previous Bag'              
  --    GoTo ErrHndlr              
  --      end              

		        select @Row= Count('x') from inserted Ins where              
    Exists (Select 'x' from Txnd Td              
      join Txnd TdRm on TdRm.TdCoCd= Td.TdCoCd              
       and TdRm.TdBYy= Td.TdBYy and TdRm.TdBChr= Td.TdBChr              
       and TdRm.TdBNo= Td.TdBNo and TdRm.TdSrNo<> 0              
      where Td.TdSrNo= 0 and Td.TdBDc= 'C' and Td.TdCoCd= Ins.TdCoCd and              
       Td.TdTc= Ins.TdTc and Td.TdYy= Ins.TdYy and              
       Td.TdChr= Ins.TdChr and Td.TdNo= Ins.TdNo and              
       TdRm.TdRmCd Not In(select OrRmCd from Bag              
        join OrdRm on BCoCd= OrCoCd and              
          BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr              
         and BOdNo= OrNo and BOdSr= OrSr              
         where BCoCd= Ins.TdCoCd and BYy= Ins.TdBYy              
          and BChr= Ins.TdBChr and BNo= Ins.TdBNo)              
         )              
          and Ins.TdSrNo= 0 and Ins.TdBDc= 'D'              
    If @Row> 0              
        Begin              
    Select @ErrMsg= 'RM Code Of Parent Not Same As Child'              
      GoTo ErrHndlr              
        end              

end
 if @wenttyp='TxnBag' 
 begin
  declare @wBLoc varchar(8)=null
  declare @wBRecvDt date=null
  declare @wBMaxTxnDt date=null
  declare @wBQty float 
  declare @wBPtQty int 
  declare @wBGrWt float 
  declare @wBRepQty float 
  declare @wBRjQty float 
  declare @wBRjPtQty int 
  declare @wBRepGrWt float
  declare @wBWrk varchar(8)

 
  If @wTcTyp<> 'BM' and (@mode= 'D' or (@mode= 'C' and (Update (TdBQty) or update (TdPtQty) or update (TdRjQty) or update (TdRjPtQty) ) ) )              
  Begin              
		select @wBQty = (case when TTrfQty=0 then (del.TdBQty* (case del.TdBDc when 'C' then 1 when 'D'              
				    then -1 else 0 end)) 
							 else
							 0 end),
		@wBRepQty= (case when TTrfQty=0 then (del.TdBQty * (case del.TdBDc when 'C' then 1 when 'D'              
					then -1 else 0 end))
							 else
							0 end),
		@wBPtQty = (case when TTrfQty=0 then (del.TdPtQty* (case del.TdBDc when 'C' then 1 when 'D'
					then -1 else 0 end)) else 0 end),              
		@wBRjQty= (case when TTrfQty=0 then (del.TdRjQty* (case del.TdRjDc when 'C' then 1 when 'D'              
					then -1 else 0 end)) else 0 end),              
		@wBRjPtQty= (case when TTrfQty=0 then (del.TdRjPtQty* (case del.TdRjDc when 'C' then 1 when 'D'              
					then -1 else 0 end)) else 0 end)             
	from deleted del join Txn on
	Txn.TIdNo=del.TdTIdNo
		where del.TdSrNo=0                
  End              

              
  If @wTcTyp= 'BM' and (@mode= 'D' or (@mode= 'C' and (Update (TdBQty) or update (TdPtQty) or update (TdRjQty) or update (TdRjPtQty) ) ) )              
  Begin              
   select 
   @wBQty= del.TdBQty * (case del.TdBDc when 'C' then 1 when 'D'              
    then 0 else 0 end),              
   @wBPtQty=  del.TdPtQty* (case del.TdBDc when 'C' then 1 when 'D'              
    then 0 else 0 end),              
   @wBRjQty= del.TdRjQty* (case del.TdRjDc when 'C' then 1 when 'D'              
    then 0 else 0 end),              
   @wBRjPtQty= del.TdRjPtQty* (case del.TdRjDc when 'C' then 1 when 'D'              
    then 0 else 0 end)              
   from deleted del join Param on PTyp= 'BCHR' and              
     PMCd= del.TdBChr and PSCd= '' and PValue in ('C', 'F') 
    and del.TdSrNo=0                
  End          
              
  select 
  @wBGrWt = (((case (case when del.TdFrRmLoc='W' then del.TdFrRmDc              
            when del.TdToRmLoc='W' then del.TdToRmDc end)              
                     when 'C' then 1 when 'D' then -1 else 0 end)* del.TdRmWt/              
   (case when tdRmCtg in ('D','C') then 5 else 1 end))),   
 
  @wBRepGrWt = (((case (case when del.TdFrRmLoc='XR' then del.TdFrRmDc              
            when del.TdToRmLoc='XR' then del.TdToRmDc end)              
                     when 'C' then 1 when 'D' then -1 else 0 end)* del.TdRmWt/              
   (case when tdRmCtg in ('D','C') then 5 else 1 end)))   
  from deleted del where del.TdSrNo<>0 and
   del.TdByy<> '' and del.TdBChr<> ''              
   and del.TdBNo<> 0 and del.TdRmCd<> ''              


 if @mode in ('D') and @wEntTyp= 'TxnBag' and (select TdSrNo from deleted)=0
 begin
  declare @wFgToBLoc varchar(8)
  declare @wTdYyKeyAckBV varchar(10)
  
  Select @wFgToBLoc = Isnull(F.FgToBLoc,'') 
	from Fgd Fd              
	join Deleted Del on FdBIdNo=Del.TdBIdNo
    join Fg F on F.FgIdNo=Fd.FdFgIdNo 
    where Fd.FdCoCd=Del.tdcocd and Del.TdSrNo=0 
     and Fd.FdBYyKey=Del.TdPrevYy + substring(str(Del.TdPrevKey),3,8)
     and FdTc='FWB' 

   Select 
	@wTdYyKeyAckBV=
	max(td.TdYyKey) 
     from Txnd td join Deleted Del
     on del.TdBIdNo=td.TdBIdNo               
     and td.TdTc='BV' and Td.TdAck='Y'
  

   Select @wBRecvDt = isnull(TDt,BOpnDt), 
		  @wBLoc= isnull(@wFgToBLoc,Txnd.TdToBLoc)
	  From Txnd 
	  join Txn on TdTIdNo=TIdNo
	  join Bag on BIdNo=Txnd.TdBIdNo 
	  join Deleted TdDel on TdDel.TdBIdNo=Txnd.TdBIdNo
	  where Txnd.TdYyKey=@wTdYyKeyAckBV 

   if @wBLoc is null
   begin
   Select @wBLoc=case when BRepYN='N' then BOpnLoc 
	else
	(select HDefRepBagLoc from Head where HCoCd=Deleted.TdCocd) end
   from Deleted, Bag where BIdNo=Deleted.TdBIdNo
   end
  End 

  If (@mode='D')
  Begin    
       select @wBMaxTxnDt = IsNull((Select Max(qDt)               
           From (Select IsNull(Max(Tddt),'01/01/80') qDt               
           From Txnd               
           where TdBIdNo=Del.TdBIdNo and TdIdNo<>del.TdIdNo                
        Union All              
                Select IsNull(Max(Fddt),'01/01/80') qDt
           From Fgd
           where FdBIdNo=Del.TdBIdNo 
		   Union All              
            Select BOpnDt qDt From Bag               
           Where BIdNo=Del.TdBIdNo ) B),null)              
   From	 Deleted Del              


  End

   Update Bag set BLoc=case when BRepYN='N' and @wBLoc is not null then @wBLoc else BLoc  end,  
				  BRecvDt = isnull(@wBRecvDt ,BRecvDt), 
				  BRepLoc = case when BRepYN='Y' and @wBLoc is not null then @wBLoc else BRepLoc end,
				  BWrk = (case when deleted.TdSrNo=0 then isnull(Txnd.TdWrk,'ZSELF') else BWrk end),
				  BLstYy =(case when @mode='D' and deleted.TdSrNo=0 then isnull(Txnd.TdYy,'') else BLstYy end),
				  BLstKey=(case when @mode='D' and deleted.TdSrNo=0 then isnull(Txnd.TdKey,0) else BLstKey end),
 				  BLstYyKey=(case when @mode='D' and deleted.TdSrNo=0 then isnull(Txnd.TdYyKey,'') else BLstYyKey end),				  
				  BMaxTxnDt = (case when deleted.TdSrNo=0 then isnull(@wBMaxTxnDt ,BOpnDt) else  BMaxTxnDt end),
				  BQty=BQty+ (case when deleted.TdSrNo=0 then case when BRepYN='N' then isnull(@wBQty,0) else 0 end else 0 end),
				  BRepQty=BRepQty + (case when deleted.TdSrNo=0 then case when BRepYN='Y' then isnull(@wBRepQty,0) else 0 end else 0 end),
				  BPtQty = BPtQty+(case when deleted.TdSrNo=0 then isnull(@wbptqty,0) else 0 end),
				  bRjQty=bRjQty+isnull(@wbRjQty,0),
				  bRjPtQty=bRjPtQty+isnull(@wBRjPtQty,0)
		from Bag 
		join Deleted on deleted.TdBIdNo=Bag.BIdNo
		left outer join Txnd on 
		deleted.TdBIdNo=Bag.BIdNo and Txnd.TdCoCd=deleted.TdCoCd and Txnd.TdYy = deleted.TdPrevYy and Txnd.TdKey=deleted.TdPrevKey 
		where Deleted.TdSrNo=0 


  Update Bag set              
  BGrWt= BGrWt+ (case when BRepYN ='N' then IsNull((Select sum((case (case when del.TdFrRmLoc='W' then del.TdFrRmDc              
            when del.TdToRmLoc='W' then del.TdToRmDc end)              
                     when 'C' then 1 when 'D' then -1 else 0 end)* del.TdRmWt/              
   (case when RmCtg in ('D','C') then 5 else 1 end))              
   from deleted del Join RmMst on RmCd= del.TdRmCd 
   where del.TdBIdNo=BIdNo and del.TdRmCd<> ''), 0) else 0 end),
  BRepGrWt= BRepGrWt+ (case when BRepYN ='Y' then IsNull((Select sum((case (case when del.TdFrRmLoc='XR' then del.TdFrRmDc              
            when del.TdToRmLoc='XR' then del.TdToRmDc end)              
                     when 'C' then 1 when 'D' then -1 else 0 end)* del.TdRmWt/              
   (case when RmCtg in ('D','C') then 5 else 1 end))              
   from deleted del Join RmMst on RmCd= del.TdRmCd where del.TdBIdNo=BIdNo and del.TdRmCd<> ''), 0) else 0 end)
  from deleted where BIdNo=deleted.TdBIdNo and deleted.TdRmCd<> ''              


  If @wEntTyp= 'TxnBag' and @wTcTyp= 'BM'              
  Begin              
   If @Mode in ('D')
   Begin              
    Update Bag               
     Set BCls = 'N',              
         BClsDt = '01/01/80'              
    From Deleted Del              
    Where BIdNo=Del.TdBIdNo and BCoCd= Del.TdCoCd And BYy= Del.TdBYy And BChr= Del.TdBChr And BNo= Del.TdBNo               
--     And Convert(Decimal(16, 3), BGrWt) <> Convert(Decimal(16, 3), 0)               
     And (Select PValue From Param Where PTyp= 'BCHR' And PMCd= Del.TdBChr And PSCd= '') in ('C', 'F')               
   End              
  End        
  
  If @wEntTyp= 'TxnBag' and @wTcTyp= 'MLT'              
  Begin              
   If @Mode in ( 'D' ,'C')             
   Begin              
    Update Bag               
     Set BCls = 'N',              
         BClsDt = '01/01/80'              
    From Deleted Del              
    Where BIdNo=Del.TdBIdNo and BCoCd= Del.TdCoCd And BYy= Del.TdBYy And BChr= Del.TdBChr And BNo= Del.TdBNo               
--     And Convert(Decimal(16, 3), BGrWt) <> Convert(Decimal(16, 3), 0)               
   End              
  End              

 End              

        
 if @mode in ('A','C') and @wEntTyp= 'TxnBag'              
 begin              
  If (@mode in( 'A','C'))              
  Begin              
   Select @wBMaxTxnDt = TdDt              
   From Inserted                  
    Join Bag on BIdNo=TdBIdNo 
   Where TdDt>=BMaxTxnDt                 
  End              
              
  select @wBWrk = case when ins.TdWrk<>'' and ins.TdSrNo=0 then ins.TdWrk else null end
  from inserted ins 

  If @wTcTyp<> 'BM' and (@mode= 'A' or (@mode= 'C' and (Update (TdBQty) or update (TdPtQty) or update (TdRjQty) or update (TdRjPtQty) or update(TdYy) or update (TdKey) or update (TdYyKey) ) ) )              
  Begin              
   select 
   @wBQty= case when TTrfQty=0 then (ins.TdBQty* (case ins.TdBDc when 'C' then -1 when 'D'              
    then 1 else 0 end)) else 0 end,              
   --BRepQty= BRepQty+ case when BRepYN ='Y' then (ins.TdBQty* (case ins.TdBDc when 'C' then -1 when 'D'              
   -- then 1 else 0 end)),              
   @wBPtQty =  case when TTrfQty=0 then (ins.TdPtQty* (case ins.TdBDc when 'C' then -1 when 'D'              
    then 1 else 0 end)) else 0 end,              
   @wBRjQty= case when TTrfQty=0 then (ins.TdRjQty* (case ins.TdRjDc when 'C' then -1 when 'D'              
    then 1 else 0 end)) else 0 end ,              
   @wBRjPtQty= case when TTrfQty=0 then (ins.TdRjPtQty* (case ins.TdRjDc when 'C' then -1 when 'D'              
    then 1 else 0 end)) else 0 end
 from inserted ins join Txn on Txn.TIdNo=ins.TdTIdNo
	join Param on PTyp= 'BCHR' and              
      PMCd= ins.TdBChr and PSCd= '' 
   where ins.TdSrNo=0         

  End              
              
  If @wTcTyp= 'BM'               
  Begin              
   If (@mode= 'A' or (@mode= 'C' and (Update (TdBQty) or update (TdPtQty) or update (TdRjQty) or update (TdRjPtQty)  ) ) )              
   Begin              
    select @wBQty=
     ((ins.TdBQty* (case ins.TdBDc when 'C' then -1 when 'D'              
     then 0 else 0 end))) ,              
    @wBPtQty=((ins.TdPtQty* (case ins.TdBDc when 'C' then -1 when 'D'              
     then 0 else 0 end)))              
    from inserted ins where ins.TdSrNo= 0 and (Select PValue from Param where PTyp= 'BCHR' and              
      PMCd= ins.TdBChr and PSCd= '' ) in ('C', 'F')               
   End               
                 
  End              
              
  Select @wBLoc= ins.TdToBLoc ,              
   @wBRecvDt=ins.TdDt 
   from inserted ins where 
    ins.TdToBLoc<> '' and ins.TdSrNo= 0 and ins.TdAck= 'Y'              
              
 
   update Bag set 
	   BMaxTxnDt = @wBMaxTxnDt,
	   BLoc=case when BRepYN='N' and @wBLoc is not null then @wBLoc else BLoc end,
	   BRecvDt=isnull(@wBRecvDt, BRecvDt),
	   --5716 repair bv issue- BRepLoc=case when BRepYn='N' and @wBLoc is not null then @wBLoc else BRepLoc end,
	   BRepLoc=case when BRepYn='Y' and @wBLoc is not null then @wBLoc else BRepLoc end,
	   BWrk=isnull(@wBWrk,Bwrk),
       BLstYy= TdYy ,              
       BLstKey= convert(int, substring((ins.TdYyKey),3, 8))              ,
       BLstYyKey= (ins.TdYyKey), 
	   BQty=Bqty+case when BRepYN='N' then isnull(@wBQty,0) else 0 end,
	   BRepQty=BRepQty+case when BRepYN='N' then isnull(@wBRepQty,0) else 0 end,
	   BGrWt =BGrWt+case when BRepYN='N' then isnull(@wBGrWt,0) else 0 end,
	   BPtQty=BPtQty+case when BRepYN='N' then isnull(@wBPtQty,0) else 0 end,
	   BRepGrWt =BRepGrWt+case when BRepYN='Y' then isnull(@wBRepGrWt,0) else 0 end,
	   BRjQty=bRjQty+isnull(@wBRjQty,0),
       BRjPtQty=bRjPtQty+isnull(@wBRjPtQty,0)
   from inserted ins where BIdNo=ins.TdBIdNo          and ins.tdsrno=0


  update Bag set              
  BMaxTxnDt=isnull(@wBMaxTxnDt,BOpnDt),
  BGrWt= BGrWt+ (case when BRepYN='N' then IsNull((Select ((case (case when ins.TdFrRmLoc='W' then ins.TdFrRmDc              
          when ins.TdToRmLoc='W' then ins.TdToRmDc end)              
          when 'C' then -1 when 'D' then +1 else 0 end)              
   * ins.TdRmWt/ (case when RmCtg in ('D','C') then 5 else 1 end))              
   from inserted ins Join RmMst on RmCd= ins.TdRmCd		
   where ins.TdBIdNo=BIdNo and ins.TdCoCd= BCoCd and ins.TdBYy= BYy and ins.TdBChr= BChr              
   and ins.TdBNo= BNo and ins.TdRmCd<> ''), 0) else 0 end),             
  BRepGrWt= BRepGrWt+ (case when BRepYN='Y' then IsNull((Select ((case (case when ins.TdFrRmLoc='XR' then ins.TdFrRmDc              
          when ins.TdToRmLoc='XR' then ins.TdToRmDc end)              
          when 'C' then -1 when 'D' then +1 else 0 end)              
   * ins.TdRmWt/ (case when RmCtg in ('D','C') then 5 else 1 end))              
   from inserted ins Join RmMst on RmCd= ins.TdRmCd 
   where ins.TdBIdNo=BIdNo and ins.TdRmCd<> ''), 0) else 0 end)
  from inserted where BIdNo=inserted.TdBIdNo and inserted.TdRmCd<> ''              

 end              
 
 IF @wEntTyp='TxnBag'              
 Begin              
  Select @Row= Count('x') From Inserted Ins Where 
   Not Exists (Select 'x' From Bag Where BIdNo=Ins.TdBIdNo and BCoCd= Ins.TdCoCd and BYy= Ins.TdByy              
     and BChr= Ins.TdBChr and BNo= Ins.TdBNo and (BCls= 'N' or BRepYN ='Y'))              
     If @Row>0              
  begin              
   Select @ErrMsg='Invalid Bag No/Bag Is Closed'              
     GoTo ErrHndlr              
  end              
  --5717
    select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and @wTcTyp              
   in('DT', 'BV', 'MLT') and TdPrtKey=@wCurrPrtn and TdSrNo= 0 and not exists(Select 'x' from Txn where              
    TPrtKey=TdPrtKey and TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo and              
   TFrBLoc= (case when TdBDc= 'C' then TdFrBLoc when TdBDc= 'D' then TdToBLoc else TdFrBLoc end))              
   and IsNull((Select hOthLocBag from Head where HCoCd= TdCoCd and HCd= @wHCd), '') = 'N'              
    if @row> 0              
  begin              
      select @errmsg= 'From Bag Loc Should Be Equal To The Default Bag Location'              
   goto ErrHndlr              
    end 

      select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and @wTcTyp               
   in('BS', 'BM') and TdPrtKey=@wCurrPrtn and TdSrNo= 0 and not exists(Select 'x' from Txn where               
    TPrtKey=TdPrtKey and TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr               
    and TNo= TdNo and TFrBLoc= TdFrBLoc)               
   and IsNull((Select hOthLocBag from Head where HCoCd= TdCoCd and HCd= @wHCd), '') = 'N'               
    if @row> 0               
  begin               
      select @errmsg= 'From Bag Loc Should Be Equal To The Default Bag Location'               
   goto ErrHndlr               
    end              
  /* **** Zubin 211 **** */              
  
     Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPrtn and              
   Not Exists (Select 'x' From Param where PTyp='DC'              
        and PMCd = Ins.TdBDc and PScd='')              
   and TdBDc<>''              
    If @Row>0              
  Begin              
   Select @ErrMsg='Invalid Bag Dc'              
   GoTo ErrHndlr              
  End              
  select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and @wTcTyp               
   in('BM') and TdPrtKey=@wCurrPrtn and ((Select count(distinct TdBChr) from Txnd where TdCoCd = inserted.TdCoCd and TdTc = inserted.TdTc and TdYy = Inserted.TdYy and TdChr = inserted.TdChr 
									and TdNo = inserted.TdNo and TdSrNo = 0 and TdBDc='D' and tdrefkey = inserted.tdrefkey)>1 
									or (Select count(distinct TdBChr) from Txnd where TdCoCd = inserted.TdCoCd and TdTc = inserted.TdTc and TdYy = Inserted.TdYy and TdChr = inserted.TdChr 
									and TdNo = inserted.TdNo and TdSrNo = 0 and TdBDc='C' and tdrefkey = inserted.tdrefkey)>1)  
                
    if @row> 0               
  begin               
      select @errmsg= 'Issue CHR must be one Receive CHR must be one'               
   goto ErrHndlr               
    end              
	
   select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and @wTcTyp               
   in('BM') and TdPrtKey=@wCurrPrtn and (Select max(distinct TdBChr) from Txnd where TdCoCd = inserted.TdCoCd and TdTc = inserted.TdTc and TdYy = Inserted.TdYy and TdChr = inserted.TdChr 
									and TdNo = inserted.TdNo and TdSrNo = 0 and TdBDc='D' and tdrefkey=inserted.tdrefkey)=
									(Select max(distinct TdBChr) from Txnd where TdCoCd = inserted.TdCoCd and TdTc = inserted.TdTc and TdYy = Inserted.TdYy and TdChr = inserted.TdChr 
									and TdNo = inserted.TdNo and TdSrNo = 0 and TdBDc='C' and tdrefkey=inserted.tdrefkey) 
                
    if @row> 0               
  begin               
      select @errmsg= 'Issue CHR must be different from Recv Chr'               
   goto ErrHndlr               
    end              

  /*********** TdBGrWt >=0 **********/              
  /* Current Partition */              
  Select @Row=Count('x') From Inserted Ins Where TdBGrWt< 0              
  If @Row>0              
    Begin              
   Select @ErrMsg='Gross Wt cannot be -ve'              
     GoTo ErrHndlr              
    End              
              
--/*  ######################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ######################  */              
  Select @Row=Count('x') From Inserted Ins              
   Join Bag  On BIdNo=TdBIdNo 
  Where (TdSrNo<> 0 and Not Exists (Select 'x' from OrdRm where              
	OrOdIdNo=BOdIdNo  and OrRmCd= TdRmCd))              
   OR              
   (TdSrNo=0 and TdRmCd<> '')              
    If @Row> 0              
  BEGIN              
   Select @ErrMsg='Rm Code does not exist in Order/Rm Code Should Be Blank'              
   GoTo ErrHndlr              
  END              
 
  --SELECT @Row= COUNT('x') FROM Inserted Ins WHERE (TdSrNo<>0 AND              
  -- NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
  --    LocTyp<> 'P' and LocTyp <>'R' AND LocCd = Ins.TdToRmLoc AND LocValidYN='Y' And              
  --     (@wTcTyp NOT IN('Rec','Cnv','Bs', 'BM') OR LocTyp= 'Xc')))              
  -- OR              
  --  (TdSrNo=0 AND TdToRmLoc<>'')              
  --  IF @Row> 0              
  --BEGIN              
  -- SELECT @ErrMsg= 'Invalid To Rm Loc/ To Rm Loc Should Be Blank'              
  -- GOTO ErrHndlr              
  --END              
               
  select @row= count('x') from inserted Ins where TdSrNo<>0 and              
    Ins.TdFrRmLoc = Ins.TdToRmLoc and (Ins.TdFrRmLoc <> '' or              
    Ins.TdToRmLoc <> '')              
    if @row>0            
    begin              
      select @errmsg= 'To Rm Loc and From Rm Loc cannot be same'              
      goto ErrHndlr              
    end              

   if @wTcTyp='BS'
   begin
	  Select @Row= Count('x') from inserted Ins where TdSrNo= 0 and TdBDc= 'D' and               
       Not Exists (Select 'x' from Txnd                
      where Txnd.TdTIdNo=Ins.TdTIdNo              
		and         Txnd.TdSrno= 0 and Txnd.TdBDc= 'C')               
     If @Row > 0               
     Begin              
    Select @ErrMsg= 'No Parent Bag For The Child Bag '              
      GoTo ErrHndlr              
     end          
	    select @Row= Count('x') from inserted Ins       
  Join Txn On TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where TMultiIr='R' and       
    (Select count('x') from Txnd  where 
     Txnd.TdTIdNo=Ins.TdTIdNo and              
     Txnd.TdSrNo= 0 and Txnd.TdBDc='C')> 1              
    and TdBDc='C' and TdSrNo= 0              
   If @Row > 0              
      begin              
	 Select @ErrMsg='Only One Parent Bag allowed in Bag Split'              
    GoTo ErrHndlr              
   end              
    
    select @Row= Count('x') from inserted Ins       
  Join Txn On TIdNo=Ins.TdTIdNo   where 
    Not Exists (Select 'x' from Txnd where Txnd.TdCoCd= Ins.TdCoCd and            
         ((TMultiIr='R' and Txnd.TdYy= Ins.TdRefYy and Txnd.TdKey= Ins.TdRefKey)      
     or       
         (TMultiIr='I' and Txnd.TdCoCd=Ins.TdCoCd and Txnd.TdRefYy= Ins.TdYy and Txnd.TdRefKey= Ins.TdKey))      
         and (select BOdDmCd from Bag where BIdNo=Txnd.TdBIdNo              
       )              
           = (select BOdDmCd from Bag where BIdNo=Ins.TdBIdNo              
       )              
            )              
          and TdBDc= 'D' and TdSrNo= 0              
    and Str(TtrfQty, 16, 1) = Str(0, 16, 1)              
        If @Row > 0              
        Begin              
    Select @ErrMsg='Design Code Of Parent Not Same As Child '      
      GoTo ErrHndlr              
        end             
       

 If Exists(Select 'x' from Txn      
   Join Inserted on TdCoCd=TCoCd and TTc=TdTc and TYy=TdYy and TChr=TdChr and TNo=TdNo and       
    TIdNo=TdTIdNo 
   Where TdBDc= 'D' and TdSrNo= 0 and TMultiIr='I') and @mode='A'      
 Begin           
    select @Row= Count('x') from inserted Ins       
  where 
    Exists (Select 'x' from Txnd Td               
 join Txnd Tdr on Tdr.TdBIdNo=Td.TdBIdNo              
         and Tdr.TdSrNo<> 0               
     where Td.TdSrNo= 0 and Td.TdCoCd= Ins.TdCoCd and           
   Td.TdRefYy= Ins.TdYy and Td.TdRefKey= Ins.TdKey      
         and Tdr.TdRmCd Not in (select OrRmCd from Bag               
        join OrdRm on BOdIdNo=OrOdIdNo 
         where BIdNo=Ins.TdBIdNo)               
         )              
          and Ins.TdBDc= 'D' and Ins.TdSrNo= 0              
    If @Row > 0              
        Begin              
    Select @ErrMsg= 'RM Code Of Parent Not Same As Child'              
      GoTo ErrHndlr              
        end              
 END      
      
  /***** Bef 3.8.0 *****/      
 If Exists(Select 'x' from Txn      
   Join Inserted on TdCoCd=TCoCd and TTc=TdTc and TYy=TdYy and TChr=TdChr and TNo=TdNo and       
    TIdNo=TdTIdNo 
   Where TdBDc= 'D' and TdSrNo= 0 and TMultiIr='R')  and @mode='A'      
 Begin      
    select @Row= Count('x') from inserted Ins where 
    Exists (Select 'x' from Txnd Td               
       join Txnd Tdr on Tdr.TdBIdNo=Td.TdBIdNo              
				and Tdr.TdSrNo<> 0               
     where Td.TdSrNo= 0 and Td.TdCoCd= Ins.TdCoCd and               
         Td.TdYy= Ins.TdRefYy and Td.TdKey= Ins.TdRefKey               
         and Tdr.TdRmCd not in(select OrRmCd from Bag               
        join OrdRm on BOdIdNo=OrOdIdNo 
         where BIdNo=Ins.TdBIdNo)               
         )              
          and Ins.TdBDc= 'D' and Ins.TdSrNo= 0              
    If @Row > 0              
        Begin              
    Select @ErrMsg= 'RM Code Of Parent Not Same As Child'              
      GoTo ErrHndlr              
        end       
	Select @Row= Count('x') from inserted Ins where 
    exists (Select 'x' from Txn where 
      TIdNo=Ins.TdTIdNo 
and         Str(TtrfQty, 16, 1) <> Str(0, 16, 1)) and              
    Exists (Select 'x' from Txnd TdD where TdD.TdSrNo<> 0 
      and TdD.TdBIdNo=Ins.TdBIdNo 
    and TdD.TdRmCd not in(select OrRmCd from Txnd TdC              
          join Bag on BIdNo=TdC.TdBIdNo              
          join OrdRm on BOdIdNo=OrOdIdNo              
          where TdC.TdSrNo= 0 
         and TdC.TdCoCd= Ins.TdCoCd and TdC.TdYy= Ins.TdRefYy              
         and TdC.TdKey= Ins.TdRefKey)              
         )              
          and Ins.TdBDc= 'D' and Ins.TdSrNo= 0              
        If @Row > 0              
        Begin              
    Select @ErrMsg= 'RM Code Of Child Not Same As Parent'      
      GoTo ErrHndlr              
        end           

 End           
 
 select @Row= Count('x') from inserted Ins       
  Join Txn On TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where TMultiIr='I' and       
    (Select count('x') from Txnd  where 
     Txnd.TdTIdNo=Ins.TdTIdNo and              
     Txnd.TdSrNo= 0 and Txnd.TdBDc='D')> 1              
    and TdBDc='D' and TdSrNo= 0              
   If @Row > 0              
      begin              
    Select @ErrMsg='Only One Child Bag allowed in Bag Split'              
    GoTo ErrHndlr              
   end              

   
   select @Row= Count('x') from inserted Ins       
  Join Txn On TPrtKey=Ins.TdPRtKey and TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where Ins.TdPRtKey=@wCurrPrtn and TMultiIr='I' and       
    exists(Select 'x' from Txnd  where Txnd.TdPrtKey=Ins.TdPRtKey               
     and Txnd.TdTIdNo=Ins.TdTIdNo and Txnd.TdCoCd=Ins.TdCoCd and              
     Txnd.TdTc=Ins.TdTc and Txnd.TdYy=Ins.TdYy and              
     Txnd.TdChr=Ins.TdChr and Txnd.TdNo=Ins.TdNo and              
     Txnd.TdSr<Ins.TdSr and Txnd.TdSrNo= 0 and Txnd.TdBDc='D')      
    and TdBDc='C' and TdSrNo= 0              
   If @Row > 0              
      begin              
    Select @ErrMsg='Only 1 set allowed in Bag Split'              
    GoTo ErrHndlr              
   end
   end 

     select @row= count('x') from inserted Ins where TdPRtKey=@wCurrPrtn and TdSrNo<>0 and              
    Ins.TdFrRmLoc = Ins.TdToRmLoc and (Ins.TdFrRmLoc <> '' or              
    Ins.TdToRmLoc <> '')              
    if @row>0            
    begin              
      select @errmsg= 'To Rm Loc and From Rm Loc cannot be same'              
      goto ErrHndlr              
    end              
              
  /****** if 'TxnBag' and SrNo<> 0 then TdFrRmLoc & TdToRmLoc both should not be of LocTyp= 'W' **********/              
  /* Current Partition */              
  Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and  TdSrNo<>0 and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'W' and LocCd = Ins.TdFrRmLoc)              
     and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'W' and LocCd = Ins.TdToRmLoc)              
    If @Row>0              
  begin              
   Select @ErrMsg= 'Both From & To Rm Locs Should Not Be ''W'' '              
   GoTo ErrHndlr              
  end              
              
  Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and  TdSrNo<>0 and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'XR' and LocCd = Ins.TdFrRmLoc)              
     and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'XR' and LocCd = Ins.TdToRmLoc)              
    If @Row>0              
  begin              
   Select @ErrMsg= 'Both From & To Rm Locs Should Not Be ''XR'' '              
   GoTo ErrHndlr              
  end              


   if @wTcTyp='REJ'
   begin
   select @Row=Count('x') from inserted join Bag on BIdNo=TdBIdNo where 
   (cast(BQty As Decimal(8,2)) <> 0 or              
    (BPtQty <> 0              )
   ) and              
   ( cast(BRjQty As Decimal(8,2)) <> 0 or              
    (BRjPtQty)<> 0              
   )              
  If @Row > 0              
  begin              
   Select @ErrMsg= 'All the Pieces/ Part Pieces In The Bag Have To Be Rejected/ Derejected'              
   GoTo ErrHndlr              
  end   

  Select @Row=Count('x') From Inserted Ins              
   Where (Not Exists (Select 'x' From Param where PTyp='REJCD'  
    and PMCd = Ins.TdRjCd and PScd='')              
    and @wTcTyp='Rej' and TdSrNo=0)              
   OR              
    ((@wTcTyp<>'REJ' or TdSrNo<>0) and TdRjCd<>'')              
    If @Row> 0              
  BEGIN              
   Select @ErrMsg='Invalid Rej Cd/Rej Cd Should Not Be Entered'              
   GoTo ErrHndlr              
  END              
  End          
	-- doubt
  SELECT @Row= COUNT('x') FROM Inserted Ins WHERE (TdSrNo=0 AND              
   ((Ins.TdWrk <> ISNULL((SELECT Del.TdWrk FROM Deleted Del WHERE 
    Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   AND              
   NOT EXISTS (SELECT 'x' FROM vParam WHERE vPCoCd = TdCoCd              
    AND vPTyp='WORK' AND vPMCd = Ins.TdWrk AND vPValidYN = 'Y') OR TdWrk = '')))              
   OR              
   (TdSrNo<>0 AND TdWrk<>'')              
      IF @Row>0              
  BEGIN              
   SELECT @ErrMsg='Invalid Wrk/Worker Should Not Be Entered'              
   GOTO ErrHndlr              
  END              
              

  Select @Row= Count('x') from Deleted Join Bag on BOdIdNo=deleted.TdBIdNo and @wTChr ='Y' where 
    Exists (Select 'x' from 
						OrdDsg  where OrdDsg.OdDmStkCoCd=Deleted.TdCoCd and OrdDsg.OdDmStkYy = Deleted.TdBYy and OrdDsg.OdDmStkChr = Deleted.TdBChr and OrdDsg.OdDmStkNo = Deleted.TdBNo and 
								OrdDsg.OdTc = 'JRO' and (Deleted.ModDt < OrdDsg.ModDt or (Deleted.ModDt = OrdDsg.ModDt and Deleted.ModTime < OrdDsg.ModTime))
			)

    If @Row > 0              
         Begin              
		    Select @ErrMsg= 'Cannot Delete JRO transaction exists for this Bag'              
			GoTo ErrHndlr              
         End              
              
    
   Select @Row= Count('x') From Deleted Del Where Del.TdRmCd<> '' and              
    exists (Select 'a' from Txnd Td where Td.TdBIdNo=Del.TdBIdNo 
	and       Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
      and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
      and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
      and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
      and Str(Td.TdRmStkRt, 16, 3)= Str(Del.TdRmStkRt, 16, 3) having 
				round(sum(Td.TdRmWt* (              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end)),3) < 0 or round(sum(Td.TdRmQty*(              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end)),3) < 0 )
     If @Row>0              
   begin              
    Select @ErrMsg='Rm Wt/Qty Is Less Than Zero In Bag'              
    GoTo ErrHndlr              
   end              
    
   Select @Row= Count('x') From Inserted Ins Where Ins.TdRmCd<> '' and              
    exists (Select 'a' from Txnd Td where Td.TdBIdNo=Ins.TdBIdNo 
	and       Td.TdRmCd= Ins.TdRmCd and Td.TdLotNo= Ins.TdLotNo              
      and Str(Td.TdRmSz, 16, 4)= Str(Ins.TdRmSz, 16, 4)              
      and Str(Td.TdRmSz2, 16, 4)= Str(Ins.TdRmSz2, 16, 4)              
      and Str(Td.TdRmSz3, 16, 4)= Str(Ins.TdRmSz3, 16, 4)              
      and Str(Td.TdRmStkRt, 16, 3)= Str(Ins.TdRmStkRt, 16, 3) having 
				round(sum(Td.TdRmWt* (              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end)),3) < 0 or round(sum(Td.TdRmQty*(              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end)),3) < 0 )
     If @Row>0              
   begin              
    Select @ErrMsg='Rm Wt/Qty Is Less Than Zero In Bag'              
    GoTo ErrHndlr              
   end              
  
              
  If @wTcTyp= 'BS'              
  Begin              
    /* Current Partition */              
        select @Row= Count('x') from inserted              
    join Txn on TCoCd = TdCoCd and TTc = TdTc and TYy = TdYy              
    and TChr = TdChr and TNo = TdNo              
    where TtrfQty = 1 and TdSrNo=0 and
    (TdBQty = 0 or                
        Round(TdBQty, 1) > Round(IsNull((Select BQty from Bag           
        where BIdNo=TdBIdNo ), 0), 1))              
        If @Row > 0              
        Begin              
    Select @ErrMsg= 'Transfer Quantity cannot be zero or more than Bag Quantity'              
      GoTo ErrHndlr              
        end        
   if @mode in ('D')              
   begin              
              
    /****** Cannot Delete 'Bs', TdDc='C', Tdsrno= 0              
    If exists TdBDc='D', TdSrno=0 in same voucher ********/              
    /* Current Partition */               
    select @Row=Count('x') from Deleted Del join Txn on Txn.TIdNo = del.TdTIdNo where 
     Exists (
			 Select 'x' from Txnd where 
      Txnd.TdTIdNo=Del.TdTIdNo and Txnd.TdBDc= 'D' and TdSrNo= 0              
			)              
           and TdBDc= 'C' and TdSrNo= 0 and TTrfQty=0             
    If @Row > 0              
         begin              
       Select @ErrMsg='Cannot Delete Parent, Child Exists'              
       GoTo ErrHndlr              
         end              
              
   end              
              
  End              
              
              
--/* **** Zubin 211 **** */ /* **** Zubin 211 **** */ /* **** Zubin 211 **** */ /* **** Zubin 211 **** */              
  If @wTcTyp= 'BM'              
  Begin              
   /*** 18/09/2006 ***/              
   /* Current Partition */               
 --  select * from inserted where 
 --  exists (select sum(TdBQty) from Txnd ins where 
	--TdTIdNo=ins.TdTIdNo and TdRefyy=ins.TdRefYY and TdRefKey=ins.TdRefKey
	--and Txnd.TdBDC='C' and Txnd.TdSRNo=0 
	--group by TdBChr
	--having min(qqty) <> max(qqty))


   select @Row= Count('x') from               
     (Select sum(TdBQty) as qQty from Txnd Tc               
     where exists (Select 'x' from inserted ins               
      where ins.TdBDc= 'D' and ins.TdSrNo= 0               
      and ins.TdTIdNo=Tc.TdTIdNo               
      and ins.TdCoCd= Tc.TdCoCd and ins.TdTc= Tc.TdTc               
      and ins.TdYy= Tc.TdYy and ins.TdChr= Tc.TdChr               
	  and ins.TdRefYy=0 
      and ins.TdNo= Tc.TdNo and ins.TdRefKey= Tc.TdRefKey)              
      and Tc.TdSrNo= 0 and Tc.TdBDc= 'C'                
     group By TdBChr) Tbl               
     having convert (decimal (16, 1), min(qQty))<> convert (decimal (16, 1), max(qQty))               
     If @Row> 0               
        Begin              
    Select @ErrMsg= 'Total Issue Qty For All Bag Characters Do Not Match'              
      GoTo ErrHndlr              
     end              
--5717
   if @mode in ('A')              
   begin    	
    select @Row= Count('x') from inserted ins               
     where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo= 0 and ins.TdBDc= 'C'               
      and exists (Select 'x' from Txnd Bc               
      join Txnd Tc on Tc.TdPrtKey=Bc.TdPrtKey and Tc.TdTIdNo=Bc.TdTIdNo              
       and Tc.TdCoCd= Bc.TdCoCd and Tc.TdTc= Bc.TdTc and Tc.TdYy= Bc.TdYy               
       and Tc.TdChr= Bc.TdChr and Tc.TdNo= Bc.TdNo and Tc.TdSr= Bc.TdSr               
       and Tc.TdSrNo> 0               
      where ins.TdPrtKey=Bc.TdPrtKey and ins.TdTIdNo=Bc.TdTIdNo              
       and ins.TdCoCd= Bc.TdCoCd and ins.TdTc= Bc.TdTc and ins.TdYy= Bc.TdYy               
       and ins.TdChr= Bc.TdChr and ins.TdNo= Bc.TdNo and Bc.TdSrNo= 0               
       and Bc.TdRefKey< ins.TdRefKey               
      group By Bc.TdRefKey, Tc.TdRmCd, convert (decimal (12, 4), Tc.TdRmSz),               
		convert (decimal (12, 4), Tc.TdRmSz2),               convert (decimal (12, 4), Tc.TdRmSz3),               
       convert (decimal (12, 2), Tc.TdRmStkRt), Tc.TdLotNo               
      having convert (decimal (16, 3), sum(case               
        when Tc.TdFrRmLoc= 'W' and Tc.TdFrRmDc= 'C' then -Tc.TdRmWt               
        when Tc.TdFrRmLoc= 'W' and Tc.TdFrRmDc= 'D' then Tc.TdRmWt               
        when Tc.TdToRmLoc= 'W' and Tc.TdToRmDc= 'C' then -Tc.TdRmWt               
        when Tc.TdToRmLoc= 'W' and Tc.TdToRmDc= 'D' then Tc.TdRmWt               
        else 0 end)) <> convert (decimal (16, 3), 0)               
       or               
       sum(case when Tc.TdFrRmLoc= 'W' and Tc.TdFrRmDc= 'C' then -Tc.TdRmQty               
        when Tc.TdFrRmLoc= 'W' and Tc.TdFrRmDc= 'D' then Tc.TdRmQty               
        when Tc.TdToRmLoc= 'W' and Tc.TdToRmDc= 'C' then -Tc.TdRmQty               
        when Tc.TdToRmLoc= 'W' and Tc.TdToRmDc= 'D' then Tc.TdRmQty               
        else 0 end) <> 0               
      )               
    If @Row > 0              
       begin              
     Select @ErrMsg= 'Complete The RM Matching Of Previous Parent-Child Records Before Starting Another Set'              
     GoTo ErrHndlr              
    end                
   end              
	--5717
	   select @Row= Count('x') from inserted where TdPRtKEy=@wCurrPrtn and TdSrNo= 0 and TdRefYy<> ''               
       If @Row > 0              
        begin              
    Select @ErrMsg= 'RefYy Should Be Blank'              
      GoTo ErrHndlr              
     end 
	--5717
	  select @Row= Count('x') from inserted where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and TdRefKey= 0               
       If @Row > 0              
        begin              
    Select @ErrMsg= 'RefKey Cannot Be Blank'              
      GoTo ErrHndlr              
     end   
	 --5717
   If @Mode = 'C'               
   Begin                  
    select @Row=Count('x') from Deleted Del where Del.TdPrtKey=@wCurrPrtn and              
      Not Exists (Select 'x' from Inserted Ins where INs.TdPrtKey=Del.TdPrtKey               
       and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd= Del.TdCoCd               
       and Ins.TdTc= Del.TdTc and Ins.Tdyy= Del.TdYy               
       and Ins.TdChr= Del.TdChr and Ins.TdNo= Del.TdNo               
       and Ins.TdSr= Del.TdSr and Ins.TdSrNo= Del.TdSrNo               
       and Ins.TdRefKey= Del.TdRefKey)           
    If @Row > 0              
    begin              
     Select @ErrMsg='Cannot Change Ref Key'              
     GoTo ErrHndlr              
    end              
   End    
   --5717
           select @errmsg= TdBQty from inserted where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and TdBDc= 'D' and              
    Not Exists (Select 'x' from Bag where BPrtKey=TdPrtKey and BIdNo=TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and              
     BChr= TdBChr and BNo= TdBNo and              
    convert(decimal(16, 1), TdBQty)= convert(decimal(16, 1), BQty))              
    If @Row> 0              
        Begin              
    Select @ErrMsg= 'Cannot Change Quantity Of Receiving Bag'              
      GoTo ErrHndlr              
        end   

    			
	Select @Row= Count('x') From Inserted Ins Where TdSrNo= 0 and              
    Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= TdBChr and PSCd= '' and              
     PValue= 'P') and TdBDc= 'C'              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Primary Bag Will Always Be The Receiving Bag'              
    GoTo ErrHndlr              
   end              
                  
   Select @Row= Count('x') From Inserted Where TdSrNo= 0 and TdBDc= 'C' and TdBQty< 0              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Flute Bag Quantity Should Be >= 0 '              
    GoTo ErrHndlr              
   end   

   select @Row= Count('x') from               
    (Select sum(case when Tc.TdBDc= 'C' then Tc.TdBQty else 0 end) as qCrQty,               
      sum(case when Tc.TdBDc= 'D' then Tc.TdBQty else 0 end) as qDrQty               
     from Txnd Tc where exists(Select 'x' from inserted ins               
      where ins.TdBDc= 'D' and ins.TdSrNo= 0               
       and ins.TdTIdNo=Tc.TdTIdNo              
			and ins.TdRefYy=0 and ins.TdRefKey= Tc.TdRefKey)         
      and Tc.TdSrNo= 0               
     group By Tc.TdBChr) Tbl               
     having convert( decimal (16, 1), sum(qDrQty))> convert( decimal (16, 1), max(qCrQty))               

   If @Row > 0               
      begin              
    Select @ErrMsg= 'Receipt Quantity Cannot Be More Than Issue'              
    GoTo ErrHndlr              
   end              
              
	   select @Row= Count('x') from inserted Ins where TdPrtKEy=@wCurrPrtn and              
    TdSrNo= 0 and exists               
    (Select 'x' from Txnd Td where Td.TdPrtKEy=Ins.TdPrtKey and Td.TdTIdNo=Ins.TdTIdNo              
     and Td.TdCoCd= Ins.TdCoCd and Td.TdTc= Ins.TdTc               
     and Td.TdYy= Ins.TdYy and Td.TdChr= Ins.TdChr and Td.TdNo= Ins.TdNo               
     and Td.TdSrNo= 0 and Td.TdSr< Ins.TdSr and Td.TdRefKey> Ins.TdRefKey)               
   If @Row > 0               
      begin              
    Select @ErrMsg= 'RefKey Should Be >= That Of Previous Sr'              
    GoTo ErrHndlr              
   end              
              
			  
    /*** (Jen 2.13) ***/              
    /****** Cannot Delete 'BM', TdDc= 'C', TdSrNo= 0               
    If exists TdBDc= 'D', TdSrNo= 0, with same TdRefKey in same voucher ********/               
    /* Current Partition */              
    Select @Row= Count('x') from Deleted Del where 
     Exists (Select 'x' from Txnd Td where 
      Td.TdTIdNo=Del.TdTIdNo and Td.TdCoCd= Del.TdCoCd and Td.TdBDc= 'D' and Td.TdSrNo= 0 and               Td.TdRefYy=0 and
      Td.TdRefKey= Del.TdRefKey               
      /* and Td.TdSr> Del.TdSr */              
             )               
           and TdBDc= 'C' and TdSrNo= 0               

    If @Row > 0              
         Begin              
		    Select @ErrMsg= 'Cannot Delete Parent, Child Exists'              
			GoTo ErrHndlr              
         End              


--   End              
--  End
 End              
End              
 IF @wEntTyp='TxnRm'              
 BEGIN              
  --5717
  Select @Row=Count('x') From Inserted Ins where Ins.TdPrtKey=@wCurrPrtn and               
   TdBYy<> '' or TdBChr<>'' or TdBNo<>0 or TdBDc <> ''              
        or TdFrBLoc <> '' or TdToBLoc <> '' or TdBGrWt <> 0              
        or TdBQty <> 0 or TdPtQty <> 0              
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Bag related fields should be empty'              
   GoTo ErrHndlr              
  END     

  Select @Row=Count('x') From Inserted Ins              
   Where Not Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
    LocTyp<> 'P' and LocCd = Ins.TdFrRmLoc And LocValidYN='Y')              
   /* AND TdFrRmDc = 'C' */              
      If @Row>0              
  BEGIN              
              
	Select @ErrMsg='Invalid From Rm Loc'              
     GoTo ErrHndlr              
  END              
  --5717
    SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and              
   (Ins.TdFrRmLoc <> ISNULL((SELECT Del.TdFrRmLoc FROM Deleted Del              
    WHERE Ins.TdPrtKey=Del.TdPrtKey and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   OR              
   Ins.TdFrRmDc <> ISNULL((SELECT Del.TdFrRmDc FROM Deleted Del              
    WHERE Ins.TdPrtKey=Del.TdPrtKey and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   OR Ins.TdFrRmLoc ='')              
   AND NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
    LocTyp<> 'P' AND LocCd = Ins.TdFrRmLoc AND LocValidYN = 'Y')              
   AND TdFrRmDc = 'D'              
/* **** Zubin 211 (Or Condition for 'TdFrRmDc' was missing) **** */              
      IF @Row>0              
  BEGIN              
              
   SELECT @ErrMsg='Invalid From Rm Loc'              
     GOTO ErrHndlr              
  END   
  --5717
    Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Param where PTyp='DC'              
    and PMCd = Ins.TdFrRmDc and PScd='')               
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid From Rm DC'              
     GoTo ErrHndlr              
  END    
  --5717
    Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and Not Exists (Select 'x' From RmMst where RmCd=TdRmCd and RmPrtKey=TdPrtKey and RmZ='N')              
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Rm Code'              
     GoTo ErrHndlr              
  END    
              
/*  ####################  Manoj # Ver: 2.0.11 Patch 01 # Date: 17/08/2005  ####################  */              
  If @wTcTyp = 'PR'              
  Begin            
   /* Current Partition */              
   Select @Row=Count('x') From Inserted Ins              
    Join Loc On LocTyp = 'XP' And LocCoCd = Ins.TdCoCd              
    And LocCd = (Case When Ins.TdFrRmDc = 'C' Then Ins.TdFrRmLoc              
      When Ins.TdFrRmDc = 'D' Then Ins.TdToRmLoc End)              
   Where Not Exists (Select 'x' From RmMst where RmCd=TdRmCd and RmZ='N' And RmValidYN = 'Y')              
       If @Row>0              
   BEGIN              
    Select @ErrMsg='Rm Code Not Defined / InValid'              
      GoTo ErrHndlr              
   END              
  End              
--   the value from X is used - except for ValidYN check is not current
  --Select @Row=Count('x') From Inserted Ins              
  -- Where (Not Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
  --  LocTyp<> 'P' and LocCd = Ins.TdToRmLoc And LocValidYn='Y'              
  --  and (@wTcTyp Not In('Rec','Cnv','Bs') Or LocTyp= 'Xc')))              
  --    If @Row>0              
  --BEGIN              
  -- Select @ErrMsg='Invalid To Rm Loc'              
  --   GoTo ErrHndlr              
  --END              
  --5717
  Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
    LocTyp<> 'P' and LocCd = Ins.TdToRmLoc And LocValidYn='Y'              
    and (@wTcTyp Not In('Rec','Cnv','Bs') Or LocTyp= 'Xc')))                  
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid To Rm Loc'              
     GoTo ErrHndlr              
  END    
  --5717
    SELECT @Row= COUNT('x') FROM Inserted Ins WHERE Ins.TdPrtKey=@wCurrPrtn and              
   ((              
   (Ins.TdToRmLoc <> ISNULL((SELECT Del.TdToRmLoc FROM Deleted Del              
    WHERE Ins.TdPrtKey=Del.TdPrtKey and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   OR              
   Ins.TdToRmDc <> ISNULL((SELECT Del.TdToRmDc FROM Deleted Del              
    WHERE Ins.TdPrtKey=Del.TdPrtKey and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   )              
   AND (NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
    LocTyp<> 'P' AND LocCd = Ins.TdToRmLoc AND LocValidYN = 'Y'              
    AND (@wTcTyp NOT IN('Rec','Cnv','Bs') OR LocTyp= 'Xc'))))              
   Or Ins.TdToRmLoc = '')              
   AND TdToRmDc = 'D'              
      IF @Row>0              
  BEGIN              
   SELECT @ErrMsg='Invalid To Rm Loc'              
     GOTO ErrHndlr              
  END         
  --5717
    select @row= count('x') from inserted Ins where Ins.TdPrtKey=@wCurrPrtn and /*TdSrNo<>0 and */              
    Ins.TdFrRmLoc = Ins.TdToRmLoc and (Ins.TdFrRmLoc <> '' or              
    Ins.TdToRmLoc <> '')              
    if @row>0                begin              
      select @errmsg= 'To Rm Loc and From Rm Loc cannot be same'              
      goto ErrHndlr              
    end   
	--5717
	  Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
     LocTyp= 'W' and LocCd = Ins.TdFrRmLoc)              
           or Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'W' and LocCd = Ins.TdToRmLoc)              
    If @Row>0              
  begin              
   Select @ErrMsg= 'From Or To Rm Loc Should Not Be ''W'' '              
   GoTo ErrHndlr              
  end 
  --5717
    Select @Row= Count('x') From Inserted Ins where Ins.TdPrtKey=@wCurrPrtn and              
   TdRjDc<> '' or TdRjQty<>0 or TdRjPtQty<> 0 or TdRjCd<> ''              
        or TdWrk<> '' or TdPrdYN<> '' or TdByLoc<> ''              
   or TdByWrk<> ''              
      If @Row> 0              
  begin              
   Select @ErrMsg= 'Rej, Prd Related Fields Should Be Empty'              
     GoTo ErrHndlr              
  end 
  --5717
     select @Row= Count('x') from Inserted Ins where Ins.TdPrtKey=@wCurrPRtn and @wTc='PR' AND @WTc='SA'      
  and Not Exists(Select * from param Where Ptyp='LABQW' and PMCd=Ins.TdRtByQW)      
   If @Row> 0              
   begin              
    Select @ErrMsg= 'Invalid Rm Qw Option'              
    GoTo ErrHndlr              
   end 
   --5717
      select @Row= Count('x') from Inserted Ins where Ins.TdPrtKey=@wCurrPRtn and @wTc='PR'       
  and Ins.TdRtByQW='Q' and (TdRmCtg<>'X' and TdRmCtg<>'D' and TdRmCtg<>'C')   
   If @Row> 0              
   begin              
    Select @ErrMsg= '"Q" Allowed in Rate By QW Option Only For Dia/CS/Accessories'              
    GoTo ErrHndlr              
   end         


END              
              
              

              
 
              
              
              
              
            
 if @mode in ('C','D') /*and  @AckYN='Y'            */
 begin              
  /* Current Partition */      
              
  update SYyMm set              
  SRmCrQty= SRmCrQty- isnull((select isnull((Del.TdRmQty), 0) from Deleted Del              
     where Del.TdFrRmDc='C' and Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd              
   and Del.TdLotNo= SLotNo and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))= SYyMm and Del.TdAck='Y' ),0),              
  SRmCrWt= SRmCrWt-isnull((select isnull((Del.TdRmWt),0) from Deleted Del              
   where Del.TdFrRmDc='C' and Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd              
     and Del.TdLotNo= SLotNo and Del.TDRmSz= SRmSz and Del.TDRmSz2= SRmSz2 and Del.TDRmSz3= SRmSz3 and Del.TdRmStkRt= SRmStkRt and Del.TdFrRmLoc=SLoc              
     and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck='Y'),0),              
  SRmDrQty= SRmDrQty-isnull((select isnull((Del.TdRmQty),0) from Deleted Del              
    where Del.TdFrRmDc='D' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd              
      and Del.TdLotNo= SLotNo and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdFrRmLoc=SLoc              
      and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck='Y'),0),              
  SRmDrWt= SRmDrWt-isnull((select isnull((Del.TdRmWt),0) from Deleted Del              
      where Del.TdFrRmDc='D' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd              
      and Del.TdLotNo= SLotNo and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck = 'Y' ),0) 
  from Deleted       
   where Deleted.TdCoCd= SCoCd and Deleted.TdRmCd= SRmCd and Deleted.TdLotNo= SLotNo              
   and Deleted.TDRmSz= SRmSz and Deleted.TDRmSz2= SRmSz2 and Deleted.TDRmSz3= SRmSz3 and Deleted.TdRmStkRt= SRmStkRt and Deleted.TdFrRmLoc= SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm              
   and (Deleted.TdFrRmDc='D'  or Deleted.TdFrRmDc='C') and       
   ((@Mode='C' and Exists(Select 'x' From Inserted ins where ins.TdCoCd =Deleted.TdCoCd and ins.TdTc=Deleted.TdTc and ins.TdYy=Deleted.TdYy and      
         ins.TdChr=Deleted.TdChr and ins.TdNo=Deleted.TdNo and ins.TdSr=Deleted.TdSr and ins.TdSrNo=Deleted.TdSrNo and          
      (ins.TdCoCd<>Deleted.TdCoCd Or ins.TdRmCd<>Deleted.TdRmCd Or ins.TdLotNo<>Deleted.TdLotNo      
        Or  ins.TDRmSz<>Deleted.TDRmSz or ins.TDRmSz2<>Deleted.TDRmSz2 or ins.TDRmSz3<>Deleted.TDRmSz3 or ins.TdRmStkRt<>Deleted.TdRmStkRt Or ins.TdFrRmLoc<>Deleted.TdFrRmLoc Or       
         Ins.TdRmWt<>Deleted.TdRmWt Or Ins.TdRmQty<>Deleted.TdRmQty Or  Ins.TdFrRmDC<>Deleted.TdFrRmDC Or  
         Ins.TdAck <> Deleted.TdAck or  
         convert(int,convert(varchar(6),ins.TdDt,112))<>convert(int,convert(varchar(6),Deleted.TdDt,112))))) Or @Mode='D')      
      
       
 end              

 if @mode in ('A','C')  /* and @AckYN= 'Y'           */
 begin              
  /* Current Partition */              
  insert into SYyMm(SPrtKey,SCoCd,SRmCd,SLotNo,SRmSz,SRmSz2, SRmSz3, SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)              
   select distinct TdPrtKey, TdCoCd, TdRmCd,TdLotNo,TdRmSz,TdRmSz2, TdRmSz3, TdRmStkRt,TdFrRmLoc,convert(int,convert(varchar(6),TdDt,112)),0,0,0,0              
   from inserted       
   where not exists(select 'x' from  SYyMm where SCoCd= TdCoCd and SRmCd=TdRmCd and SLotNo= TdLotNo              
   and SRmSZ=TdRmSz and SRmSZ2=TdRmSz2 and SRmSZ3=TdRmSz3 and SRmStkRt=TdRmStkRt and SLoc=TdFrRmLoc and              
   SYyMm= convert(int,convert(varchar(6),TdDt,112))) and            TdAck='Y' and  
   ((@wEntTyp= 'TxnBag' and TdSrNo<> 0) or (@wEntTyp= 'TxnRm' and TdSrNo= 0))
   
              
  /* Current Partition */               
  update SYyMm set              
  SRmCrQty=SRmCrQty+isnull((select isnull((Ins.TdRmQty),0) from Inserted Ins       
   where Ins.TdFrRmDc='C' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdFrRmLoc=SLoc              
  and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck='Y'),0),              
  SRmCrWt=SRmCrWt+isnull((select isnull((Ins.TdRmWt),0) from Inserted Ins       
   where Ins.TdFrRmDc='C' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck='Y'),0),              
  SRmDrQty=SRmDrQty+isnull((select isnull((Ins.TdRmQty),0) from Inserted Ins       
   where Ins.TdFrRmDc='D'  and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck='Y'),0),              
  SRmDrWt=SRmDrWt+isnull((select isnull((Ins.TdRmWt),0) from Inserted Ins       
   where Ins.TdFrRmDc='D'  and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck='Y'),0)              
  from Inserted       
   where Inserted.TdCoCd= SCoCd and Inserted.TdRmCd= SRmCd and Inserted.TdLotNo=SLotNo              
   and Inserted.TDRmSz= SRmSz and Inserted.TDRmSz2= SRmSz2 and Inserted.TDRmSz3= SRmSz3  and Inserted.TdRmStkRt= SRmStkRt              
   and Inserted.TdFrRmLoc= SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm              
   and (Inserted.TdFrRmDc='D' or Inserted.TdFrRmDc='C')  and
   ((@Mode='C' and Exists(Select 'x' From Deleted Del where Inserted.TdCoCd =Del.TdCoCd and Inserted.TdTc=Del.TdTc and Inserted.TdYy=Del.TdYy and      
         Inserted.TdChr=Del.TdChr and Inserted.TdNo=Del.TdNo and Inserted.TdSr=Del.TdSr and Inserted.TdSrNo=Del.TdSrNo  and                   
      (Inserted.TdCoCd<>Del.TdCoCd Or Inserted.TdRmCd<>Del.TdRmCd Or Inserted.TdLotNo<>Del.TdLotNo      
        Or  Inserted.TDRmSz<>Del.TDRmSz or Inserted.TDRmSz2<>Del.TDRmSz2 or Inserted.TDRmSz3<>Del.TDRmSz3 or Inserted.TdRmStkRt<>Del.TdRmStkRt Or Inserted.TdFrRmLoc<>Del.TdFrRmLoc  or      
        Inserted.TdRmWt<>Del.TdRmWt Or Inserted.TdRmQty<>Del.TdRmQty Or Inserted.TdFrRmDC<>Del.TdFrRmDC Or 
          Inserted.TdAck <> Del.TdAck or      
         convert(int,convert(varchar(6),Inserted.TdDt,112))<>convert(int,convert(varchar(6),Del.TdDt,112))))) Or @Mode='A')             

end              
              
 /*** Update SYyMm from TdToRmLoc ***/              
 --select @AckYN= Del.TdAck  from deleted Del 	
 if @mode in ('C','D')  /*and @AckYN= 'Y'            */
 begin              
  /* Current Partition */              
  update SYyMm set              
  SRmCrQty=SRmCrQty-isnull((select isnull(sum(Del.TdRmQty),0) from Deleted Del       
   where Del.TdToRmDc='C' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd and Del.TdLotNo=SLotNo              
   and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and  Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck='Y'),0)              ,
  SRmCrWt=SRmCrWt-isnull((select isnull(sum(Del.TdRmWt),0) from Deleted Del       
   where Del.TdToRmDc='C' and Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd and Del.TdLotNo=SLotNo              
   and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3   and Del.TdRmStkRt=SRmStkRt and Del.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck='Y'),0)              ,
  SRmDrQty=SRmDrQty-isnull((select isnull(sum(Del.TdRmQty),0) from Deleted Del       
   where Del.TdToRmDc='D' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd and Del.TdLotNo=SLotNo              
   and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck = 'Y'),0)    ,              
  SRmDrWt=SRmDrWt-isnull((select isnull(sum(Del.TdRmWt),0) from Deleted Del       
   where Del.TdToRmDc='D' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd and Del.TdLotNo=SLotNo              
   and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck = 'Y')  ,0)            
  from Deleted       
   where Deleted.TdCoCd= SCoCd and Deleted.TdRmCd= SRmCd and Deleted.TdLotNo=SLotNo              
   and Deleted.TDRmSz= SRmSz and Deleted.TDRmSz2= SRmSz2 and Deleted.TDRmSz3= SRmSz3 and Deleted.TdRmStkRt= SRmStkRt and Deleted.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),Tddt,112))= SYyMm              
   and (Deleted.TdToRmDc='D' or Deleted.TdToRmDc='C')  and      
   ((@Mode='C' and Exists(Select 'x' From Inserted ins where ins.TdCoCd =Deleted.TdCoCd and ins.TdTc=Deleted.TdTc and ins.TdYy=Deleted.TdYy and      
         ins.TdChr=Deleted.TdChr and ins.TdNo=Deleted.TdNo and ins.TdSr=Deleted.TdSr and ins.TdSrNo=Deleted.TdSrNo and          
      (ins.TdCoCd<>Deleted.TdCoCd Or ins.TdRmCd<>Deleted.TdRmCd Or ins.TdLotNo<>Deleted.TdLotNo      
        Or  ins.TDRmSz<>Deleted.TDRmSz or ins.TDRmSz2<>Deleted.TDRmSz2 or ins.TDRmSz3<>Deleted.TDRmSz3 or    ins.TdRmStkRt<>Deleted.TdRmStkRt Or ins.TdToRmLoc<>Deleted.TdToRmLoc Or       
        Ins.TdRmWt<>Deleted.TdRmWt Or Ins.TdRmQty<>Deleted.TdRmQty Or Ins.TdToRmDC<>Deleted.TdToRmDC Or   Ins.TdAck <> Deleted.TdAck or      
         convert(int,convert(varchar(6),ins.TdDt,112))<>convert(int,convert(varchar(6),Deleted.TdDt,112))))) Or @Mode='D')      
      
             
end              
              
 --select @AckYN= Ins.TdAck  from inserted Ins 	             
 if @mode in ('A','C') /*and @AckYN = 'Y'*/
 begin              
  /* Current Partition */              
  insert into SYyMm(SPrtKey,SCoCd, SRmCd,SLotNo,SRmSz,SRmSz2, SRmSz3, SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)              
  select distinct TdPrtKey,TdCoCd, TdRmCd,TdLotNo,TdRmSz,TdRmSz2, TdRmSz3, TdRmStkRt, TdToRmLoc, convert(int,convert(varchar(6),TdDt,112)), 0, 0, 0, 0              
   from inserted       
   where not exists(select 'x' from  SYyMm where SCoCd= TdCoCd and SRmCd=TdRmCd and SLotNo=TdLotNo              
   and SRmSZ=TdRmSz and SRmStkRt=TdRmStkRt and SLoc=TdToRmLoc and    SRmSZ2=TdRmSz2 and    SRmSZ3=TdRmSz3        and       
   SYyMm=convert(int,convert(varchar(6),TdDt,112))) and TdAck = 'Y' and             
   ((@wEntTyp= 'TxnBag' and TdSrNo<> 0) or (@wEntTyp= 'TxnRm' and TdSrNo= 0))              
    
   
  /* Current Partition */              
  update SYyMm set              
  SRmCrQty=SRmCrQty+isnull((select isnull((Ins.TdRmQty),0) from Inserted Ins       
   where Ins.TdToRmDc='C' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck = 'Y'),0),              
  SRmCrWt=SRmCrWt+isnull((select isnull((Ins.TdRmWt),0) from Inserted Ins       
   where Ins.TdToRmDc= 'C' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck = 'Y'),0) ,             
  SRmDrQty=SRmDrQty+isnull((select isnull((Ins.TdRmQty),0) from Inserted Ins       
   where Ins.TdToRmDc='D' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck = 'Y'),0),              
  SRmDrWt=SRmDrWt+isnull((select isnull((Ins.TdRmWt),0) from Inserted Ins       
   where Ins.TdToRmDc='D' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck = 'Y'),0)              
  from Inserted       
   where Inserted.TdCoCd= SCoCd and Inserted.TdRmCd=SRmCd and Inserted.TdLotNo=SLotNo              
   and Inserted.TDRmSz=SRmSz and Inserted.TDRmSz2=SRmSz2 and Inserted.TDRmSz3=SRmSz3 and Inserted.TdRmStkRt=SRmStkRt and Inserted.TdToRmLoc= SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm              
   and (Inserted.TdToRmDc='D' or Inserted.TdToRmDc='C')  and      
 ((@Mode='C' and Exists(Select 'x' From Deleted Del where Inserted.TdCoCd =Del.TdCoCd and Inserted.TdTc=Del.TdTc and Inserted.TdYy=Del.TdYy and      
         Inserted.TdChr=Del.TdChr and Inserted.TdNo=Del.TdNo and Inserted.TdSr=Del.TdSr and Inserted.TdSrNo=Del.TdSrNo and                   
      (Inserted.TdCoCd<>Del.TdCoCd Or Inserted.TdRmCd<>Del.TdRmCd Or Inserted.TdLotNo<>Del.TdLotNo      
        Or  Inserted.TDRmSz<>Del.TDRmSz Or  Inserted.TDRmSz2<>Del.TDRmSz2 or  Inserted.TDRmSz3<>Del.TDRmSz3 or Inserted.TdRmStkRt<>Del.TdRmStkRt Or Inserted.TdToRmLoc<>Del.TdToRmLoc Or       
          Inserted.TdRmWt<>Del.TdRmWt Or Inserted.TdRmQty<>Del.TdRmQty Or  Inserted.TdToRmDC<>Del.TdToRmDC Or 
		  Inserted.TdAck <> Del.TdAck or  
         convert(int,convert(varchar(6),Inserted.TdDt,112))<>convert(int,convert(varchar(6),Del.TdDt,112))))) Or @Mode='A')                   
      
 end              
              
 /* urmi NegStk Start*/              
 IF @wEntTyp in ('TxnRm', 'TxnBag')              
 BEGIN
 
 if @mode in ('A','C')              
 begin              
 
  Select @Row=Count('x')               
     From Inserted Ins               
      Join Txn on TIdNo=Ins.TdTIdNo and TCoCd= Ins.TdCoCd and TTc= Ins.TdTc and TYy= ins.TdYy and              
        TChr= Ins.TdChr and TNo= Ins.TdNo               
      Join Head On HCoCd=TCoCd And HCd= @wHCd And HNegStkYn= 'N'              
      Join RmMst On RmCd=Ins.TdRmCd And               
        (Case When hNegStkRmCtg = '' Then ','+RmCtg+','               
        Else (Select ','+PDesc225+',' From Param               
            Where PTyp='GLDAS' And PMCd=hNegStkRmCtg) End) Like '%,'+RmCtg+',%'              
  Where Exists (Select 'X' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and              
        HNegStkYn= 'N')              
  and not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and               
    PMCd= Ins.ModUsr and UaMnuCd = Ins.TdTc and UaSeoYn= 'Y' and PValue= TSeoPwd)               
  and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and TSeoPwd= PValue1)               
  and Exists (Select 'X' From Loc where LocCoCd= Ins.TdCoCd and LocTyp = 'S'              
   and LocCd = (case when Ins.TdFrRmDC = 'C' then Ins.TdFrRmLoc              
                         when Ins.TdFrRmDC = 'D' then Ins.TdToRmLoc              
          else '' end) )              
  and ((Select CAST(SUM(SRmDrWt)-SUM(SRmCrWt) As Decimal(10,3)) from SYyMm              
    where Ins.TdCoCd= SCoCd and Ins.TdRmCd= SRmCd and Ins.TdLotNo= SLotNo and              
     Ins.TDRmSz= SRmSz and              Ins.TDRmSz2= SRmSz2 and              Ins.TDRmSz3= SRmSz3 and              
     (case when Ins.TdFrRmDC = 'C' then Ins.TdFrRmLoc              
 when Ins.TdFrRmDC = 'D' then Ins.TdToRmLoc              
     else '' end )= SLoc              
       )
      ) < 0              
  And (Exists (Select 'X' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and              
       HAvgSRtYN = 'Y')              
       Or              
       Exists (Select 'X' from Param Join RmMst On PTyp = 'RMSCTG' AND RmCd = Ins.TdRmCd And              
   PMCD = RMCTG And PSCd = RMSCTG where PValue1 = 'N'))              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Only Authorized Users can enter Negative Stock For This Rm Ctg'              
     GoTo ErrHndlr              
       end              
              
  Select @Row=Count('x') From Inserted Ins               
    join Txn on TIdNo=Ins.TdTIdNo and TCoCd= Ins.TdCoCd and              
       TTc= Ins.TdTc and TYy= ins.TdYy and TChr= Ins.TdChr and TNo= Ins.TdNo              
      Join Head On HCoCd=TCoCd And HCd= @wHCd And HNegStkYn= 'N'              
      Join RmMst On RmCd=Ins.TdRmCd And               
        (Case When hNegStkRmCtg = '' Then ','+RmCtg+','               
    Else (Select ','+PDesc225+',' From Param               
            Where PTyp='GLDAS' And PMCd=hNegStkRmCtg) End) Like '%,'+RmCtg+',%'              
  Where Exists (Select 'X' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and              
        HNegStkYn= 'N')              
  and not Exists (select 'x' from Param Join UsrAccess Ua on Ua.UaUsrCd = PMCd where PTyp= 'USR' and               
    PMCd= ins.ModUsr and Ua.UaMnuCd = ins.TdTc and Ua.UaSeoYn= 'Y' and PValue= TSeoPwd)               
  and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and TSeoPwd= PValue1)               
  and Exists (Select 'X' From Loc where LocCoCd= Ins.TdCoCd and LocTyp = 'S'              
   and LocCd = (case when Ins.TdFrRmDC = 'C' then Ins.TdFrRmLoc              
                         when Ins.TdFrRmDC = 'D' then Ins.TdToRmLoc              
          else '' end) )              
              
  and ( (Select CAST(SUM(SRmDrWt)-SUM(SRmCrWt) As Decimal(10,3)) from SYyMm              
    where Ins.TdCoCd= SCoCd and Ins.TdRmCd= SRmCd and Ins.TdLotNo= SLotNo and              
     Ins.TDRmSz= SRmSz and Ins.TDRmSz2= SRmSz2 and Ins.TDRmSz3= SRmSz3 and Ins.TdRmStkRt =SRmStkRt and              
     (case when Ins.TdFrRmDC = 'C' then Ins.TdFrRmLoc              
        when Ins.TdFrRmDC = 'D' then Ins.TdToRmLoc              
      else '' end )= SLoc              
     )/*- Ins.TdRmWt */              
      ) < 0              
  And Exists (Select 'X' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and              
       HAvgSRtYN = 'N')              
  And Exists (Select 'X' from Param Join RmMst On PTyp = 'RMSCTG' AND RmCd = Ins.TdRmCd And              
   PMCD = RMCTG And PSCd = RMSCTG where PValue1 = 'Y')              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Only Authorized Users can enter Negative Stock For This Rm Ctg'              
     GoTo ErrHndlr              
       end              
              
              
      end               
      if @mode in ('D')              
      begin              
  Select @Row=Count('x') From Deleted Del join Txn on 
       TIdNo=Del.TdTIdNo and TCoCd= Del.TdCoCd and              
       TTc= Del.TdTc and TYy= Del.TdYy and              
       TChr= Del.TdChr and TNo= Del.TdNo              
      Join Head On HCoCd=TCoCd And HCd= @wHCd And HNegStkYn= 'N'              
      Join RmMst On RmCd=Del.TdRmCd And               
        (Case When hNegStkRmCtg = '' Then ','+RmCtg+','               
         Else (Select ','+PDesc225+',' From Param               
            Where PTyp='GLDAS' And PMCd=hNegStkRmCtg) End) Like '%,'+RmCtg+',%'              
  Where Exists (Select 'X' From Head where HCoCd= Del.TdCoCd and HCd= @wHCd and              
        HNegStkYn= 'N')              
  and not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and               
    PMCd= Del.ModUsr and UaMnuCd = Del.TdTc and UaSeoYn= 'Y' and PValue= TSeoPwd)               
  and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and TSeoPwd= PValue1)               
  and Exists (Select 'X' From Loc where LocCoCd= Del.TdCoCd and LocTyp = 'S'              
   and LocCd = (case when Del.TdFrRmDC = 'D' then Del.TdFrRmLoc              
                          when Del.TdFrRmDC = 'C' then Del.TdToRmLoc              
          else '' end) )              
  and ( (Select CAST(SUM(SRmDrWt)-SUM(SRmCrWt) As Decimal(10,3)) from SYyMm              
    where Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd and Del.TdLotNo= SLotNo and              
     Del.TDRmSz= SRmSz and              Del.TDRmSz2= SRmSz2 and              Del.TDRmSz3= SRmSz3 and              
     (case when Del.TdFrRmDC = 'D' then Del.TdFrRmLoc              
     when Del.TdFrRmDC = 'C' then Del.TdToRmLoc              
     else '' end )= SLoc              
       ) /*- Del.TdRmWt */              
      ) < 0              
  And (Exists (Select 'X' From Head where HCoCd= Del.TdCoCd and HCd= @wHCd and              
       HAvgSRtYN = 'Y')              
    Or              
       Exists (Select 'X' from Param Join RmMst On PTyp = 'RMSCTG' AND RmCd = Del.TdRmCd And              
   PMCD = RMCTG And PSCd = RMSCTG where PValue1 = 'N'))              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Only Authorized Users can enter Negative Stock For This Rm Ctg'              
     GoTo ErrHndlr              
  end              
              
  Select @Row=Count('x') From Deleted Del join Txn on TIdNo=Del.TdTIdNo              
       and TCoCd= Del.TdCoCd and TTc= Del.TdTc and TYy= Del.TdYy and              
       TChr= Del.TdChr and TNo= Del.TdNo              
      Join Head On HCoCd=TCoCd And HCd= @wHCd And HNegStkYn= 'N'              
      Join RmMst On RmCd=Del.TdRmCd And               
        (Case When hNegStkRmCtg = '' Then ','+RmCtg+','               
          Else (Select ','+PDesc225+',' From Param             
            Where PTyp='GLDAS' And PMCd=hNegStkRmCtg) End) Like '%,'+RmCtg+',%'              
  Where Exists (Select 'X' From Head where HCoCd= Del.TdCoCd and HCd= @wHCd and              
        HNegStkYn= 'N')              
  and not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and               
    PMCd= Del.ModUsr and UaMnuCd = Del.TdTc and UaSeoYn= 'Y' and PValue= TSeoPwd)               
  and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and TSeoPwd= PValue1)               
  and Exists (Select 'X' From Loc where LocCoCd= Del.TdCoCd and LocTyp = 'S'              
   and LocCd = (case when Del.TdFrRmDC = 'D' then Del.TdFrRmLoc              
                          when Del.TdFrRmDC = 'C' then Del.TdToRmLoc              
          else '' end) )              
              
  and ((Select CAST(SUM(SRmDrWt)-SUM(SRmCrWt) As Decimal(10,3)) from SYyMm              
    where Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd and Del.TdLotNo= SLotNo and              
     Del.TDRmSz = SRmSz and Del.TDRmSz2 = SRmSz2 and Del.TDRmSz3 = SRmSz3 and Del.TdRmStkRt = SRmStkRt and              
     (case when Del.TdFrRmDC = 'D' then Del.TdFrRmLoc              
     when Del.TdFrRmDC = 'C' then Del.TdToRmLoc              
     else '' end )= SLoc              
       ) /*- Del.TdRmWt */              
      ) < 0              
  And Exists (Select 'X' From Head where HCoCd= Del.TdCoCd and HCd= @wHCd and              
       HAvgSRtYN = 'N')              
  And Exists (Select 'X' from Param Join RmMst On PTyp = 'RMSCTG' and RmCd = Del.TdRmCd And              
       PMCD = RMCTG And PSCd = RMSCTG where PValue1 = 'Y')              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Only Authorized Users can enter Negative Stock For This Rm Ctg'              
     GoTo ErrHndlr              
  End              
      End  
	  END              
 

 /**** urmi NegStk end ****/              
              
 If @wEntTyp = 'TxnBag' And @wTcTyp= 'BM'               
 Begin              
  If @Mode in ('A', 'C')              
  Begin              
   Update Bag               
    Set BCls = 'Y',              
                      BClsDt= (Select TDt From Txn Where 
          TIdNo=ins.TdTIdNo and TCoCd= ins.TdCoCd And TTc = ins.TdTc               
          And TYy = ins.TdYy And TChr= ins.TdChr And TNo= ins.TdNo)               
   From inserted ins              
   Where BIdNo=ins.TdBIdNo and BCoCd= ins.TdCoCd And BYy= ins.TdBYy And BChr= ins.TdBChr               
    And BNo= ins.TdBNo And ins.TdSrNo= 0               
    And Convert( Decimal (16, 1), BQty) = Convert( Decimal (16, 1), 0)               
    And Convert( Decimal (16, 3), BGrWt) = Convert( Decimal (16, 3), 0)               
    And BPtQty = 0               
    And (Select PValue From Param Where PTyp= 'BCHR' And PMCd= ins.TdBChr And PSCd= '') in ('C', 'F')               
  End              
 End  
  If @wEntTyp= 'TxnBag' and @wTcTyp= 'MLT'              
  Begin              
   If @Mode= 'D'              
   Begin              
    Update Bag               
     Set BCls = 'N',              
         BClsDt = '01/01/80'              
    From Deleted Del              
    Where BIdNo=Del.TdBIdNo and BCoCd= Del.TdCoCd And BYy= Del.TdBYy And BChr= Del.TdBChr And BNo= Del.TdBNo               
--     And Convert(Decimal(16, 3), BGrWt) <> Convert(Decimal(16, 3), 0)               
   End              
  End              

 If @wEntTyp = 'TxnBag' And @wTcTyp= 'MLT'               
 Begin              

  If @Mode in ('A', 'C')              
  Begin              
   --update bag set bOpnDesc=(select Convert(Decimal (16, 3), BGrWt) from bag where bidno=(select inserted.tdbidno from inserted))
   --where BIdNo=18168

   Update Bag               
    Set BCls = 'Y',              
                      BClsDt= (Select TDt From Txn Where 
          TIdNo=ins.TdTIdNo and TCoCd= ins.TdCoCd And TTc = ins.TdTc               
          And TYy = ins.TdYy And TChr= ins.TdChr And TNo= ins.TdNo)               
   From inserted ins              
   Where BIdNo=ins.TdBIdNo and BCoCd= ins.TdCoCd And BYy= ins.TdBYy And BChr= ins.TdBChr               
    And BNo= ins.TdBNo             
    And Convert(Decimal (16, 1), BQty) = Convert( Decimal (16, 1), 0)               
    And Convert(Decimal (16, 3), BGrWt) = Convert( Decimal (16, 3), 0)               
    And BPtQty = 0               
  End              
 End  
Return              
ErrHndlr:              
 BEGIN              
  Select @ErrMsg='|'+@ErrMsg              
  RaisError(@ErrMsg,16,1)              
  Rollback Transaction              
 END 
End



GOGO



ALTER TRIGGER [dbo].[TxndPrd_INUPDL] ON [dbo].[TxndPrd] 
        FOR INSERT,UPDATE,DELETE  
		NOT FOR REPLICATION
AS  
BEGIN  
	 /* Trigger Version 3.11.0 */
	 DECLARE @ROW INT  
	 DECLARE @mode char(1)  
	 DECLARE @errmsg varchar(255)  
	 DECLARE @ReplAgt BIT  
	 DECLARE @wCurrPrtn VarChar(1)

	IF @@ROWCOUNT=0   RETURN  
  
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'TxndPrd' )  
	RETURN  
  
	select @mode=case (select count('x') from inserted) when 0 then 'D'  
       else  
	case (select count('x') from deleted) when 0 then 'A' else 'C' end  
       end   

	Select @wCurrPrtn='C'
	
	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))


	if @mode in ('A', 'C')
	Begin
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and not exists (select 'x' from Txnd where
			TdPrtKey=TpPrtKey and TdIdNo=TpTdIdNo and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr
			and TdNo= TpNo and TdSr= TpSr ) /*and TdPrdYN= 'Y'*/
		if @row>0
		begin
			select @errmsg= 'Voucher Sr Not Defined In Txnd' /* /TdPrdYn = 'N' */
			goto ErrHndlr
		end
		--5717
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and TpSrNo<= 0
		if @row>0
		begin
			select @errmsg= 'Serial Number Cannot Be Zero'
			goto ErrHndlr
		end

		/*** Valid Worker ***/
		/* Current Partition */
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and TpByWrk<> '' and
			not exists (select vPMCd from vParam where vPCoCd = TpCoCd
				and vPTyp= 'WORK' and vPMCd= TpByWrk and vPValidYN='Y')
		if @row>0
		begin
			select @errmsg= 'Worker Code Not Defined'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and not exists (select 'x' from Txnd where
			TdPrtKey=TpPRtKey and TdIdNo=TpTdIdNo and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr
			and TdNo= TpNo and TdSr= TpSr and TdBQty>= 0)
		if @row>0
		begin
			select @errmsg= 'Bag Qty Should Be >= 0'
			goto ErrHndlr
		end

		select @Row= Count('x') from inserted join Txnd on TdPrtKey=TpPrtKey and TdIdNo=TpTdIdNo
				and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr
				and TdNo= TpNo and TdSr= TpSr and TdSrNo= 0
			where TpPrtKey=@wCurrPrtn and not exists (Select 'x' from Bag where BPrtKey=TdPrtKey and BIdno=TdBIdNo 
				and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo and BLstYyKey= TdYyKey)
		If @Row > 0
		begin
			Select @ErrMsg= 'Cannot Edit Since Not The Last Transaction Of The Bag'
			GoTo ErrHndlr
		end
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and TpSubPrc<> '' and
				not exists (select PSCd from Param where PTyp= 'SUBPRC'
					and PMCd= (Select LocPrcs from Loc where LocCoCd =TpCoCd and 
					LocCd = (Select TdByLoc from Txnd where TdPRtKey=TpPrtKey and TdIdNo=TpTdIdNo 
					and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr and TdNo= TpNo
					and TdSr= TpSr and TdSrNo= 0) and PSCd= TpSubPrc))
		if @row>0
		begin
			select @errmsg= 'Sub Process Not Defined'
			goto ErrHndlr
		end
		--5717
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and (TpPrdCtg= '' or TpRmQty<> 0 or TpRmCtg<> '' or TpRmSCtg<> '' or TpSetTyp<> '') and
					(Select (Select LocForPrcs.LocPrdTyp from Loc join Loc LocForPrcs on LocForPrcs.LocCoCd = Loc.LocCoCd and LocForPrcs.LocCd =Loc.LocPrcs  where Loc.LocCoCd= TdCoCd and Loc.LocCd= TdByLoc) 
						from Txnd where TdPrtKey=TpPrtKey and TdIdNo=TpTdIdNo and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy 
						and TdChr= TpChr and TdNo= TpNo and TdSr= TpSr and TdSrNo= 0 ) in ('GN')
		if @row> 0
		begin
			select @errmsg= 'For PrdTyp GN, Rm Qty must be 0, PrdCtg cannot be Blank, RmCtg/RmSubCtg/SetTyp must be blank'
			goto ErrHndlr
		end
		--5717
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and (TpPrdCtg<> '' or TpRmQty= 0 or TpRmCtg= '' or TpRmSCtg= '' or TpSetTyp= '') and
					(Select (Select LocForPrcs.LocPrdTyp from Loc join Loc LocForPrcs on LocForPrcs.LocCoCd = Loc.LocCoCd and LocForPrcs.LocCd =Loc.LocPrcs  where Loc.LocCoCd= TdCoCd and Loc.LocCd= TdByLoc) 
					From Txnd where TdPrtKey=TpPrtKey and TdIdNo=TpTdIdNo and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr 
						and TdNo= TpNo and TdSr= TpSr and TdSrNo= 0) in ('WS', 'GS')
		if @row> 0
		begin
			select @errmsg= 'For PrdTyp WS, GS, Rm Qty cannot be 0, PrdCtg must be Blank, RmCtg/RmSubCtg/SetTyp cannot be Blank'
			goto ErrHndlr
		end
		--5717
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and (TpPrdCtg<> '' or TpRmQty= 0 or TpRmCtg= '' or TpRmSCtg= '' or TpSetTyp<> '') and 
					(Select (Select LocForPrcs.LocPrdTyp from Loc join Loc LocForPrcs on LocForPrcs.LocCoCd = Loc.LocCoCd and LocForPrcs.LocCd =Loc.LocPrcs  where Loc.LocCoCd= TdCoCd and Loc.LocCd= TdByLoc)  
					from Txnd where TdPrtKey=TpPrtKey and TdIdNo=TpTdIdNo and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr 
						and TdNo= TpNo and TdSr= TpSr and TdSrNo= 0) in ('WP', 'GP')
		if @row> 0
		begin
			select @errmsg= 'For PrdTyp GP, WP Rm Qty cannot be zero, PrdCtg must be blank, RmCtg/RmSubCtg cannot be blank, SetTyp must be blank'
			goto ErrHndlr
		end

		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and TpPrdCtg<> '' and
	  		not exists (select 'x' from Param where PTyp= 'PRDCTG' and PMCd= TpPrdCtg)
 		if @row> 0
	 	Begin
  			select @errmsg= 'Production Category Not Defined'
  			goto ErrHndlr
	 	end

		/* Current Partition */
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and TpRmCtg<> '' and not exists (select PMCd from Param where PTyp= 'RMCTG'
				and PMCd= TpRmCtg)
		if @row>0
		begin
			select @errmsg= 'Raw Material Category Not Defined'
			goto ErrHndlr
		end

		/* Current Partition */
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and TpRmSCtg<> '' and not exists (select PSCd from Param where PTyp= 'RMSCTG'
				and PMCd= TpRmCtg and PSCd= TpRmSCtg)
		if @row>0
		begin
			select @errmsg= 'Raw Material Sub Category Not Defined'
			goto ErrHndlr
		end

		Select @row= count('x') from inserted 
					Join Head On HCoCd=TpCoCd And HCd='ZSELF'
		Where TpPrtKey=@wCurrPrtn and TpSetTyp<> '' and
		 ((hSetTypOnlyFrOrdYN='N' And Not exists (Select 'x' from Param where PTyp= 'LABSCD' and PMCd= 'SET' and PSCd= TpSetTyp)) OR
		  (hSetTypOnlyFrOrdYN='Y' And Not exists (select 'x' from Txnd 												
														Join Bag on BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo
														Join OrdRm on OrCoCd= BCoCd and OrTc= BOdTc and OrYy= BOdYy and OrChr= BOdChr and OrNo= BOdNo and OrSr= BOdSr
													Where TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr and TdNo= TpNo and TdSr= TpSr 
														And TdSrNo=0 And OrRmCtg= TpRmCtg and OrRmSCtg= TpRmSCtg And OrSetSCd=TpSetTyp))) 			

 		if @row> 0
	 	Begin
  			select @errmsg= 'Setting Type Not Defined or Absent in Order'
  			goto ErrHndlr
	 	end
		--5717
		select @row= count('x') from inserted where TpPrtKey=@wCurrPrtn and convert (decimal (16, 2), TpPts)= convert (decimal (16, 2), 0)		
		if @row> 0
		begin
			select @errmsg= 'Production Points Cannot Be Zero'
			goto ErrHndlr
		end

	End
	--5717
	if @mode in ('C')
	Begin
		/****** Key Cannot Be Changed ******/		
		select @Row=Count('x') from deleted del where del.TpPrtKey=@wCurrPrtn and
			not exists (Select 'x' from inserted ins where ins.TpPrtKey=del.TpPrtKey and ins.TpIdNo=del.TpIdNo
			and ins.TpCoCd= del.TpCoCd and ins.TpTc= del.TpTc and ins.TpYy= del.TpYy and ins.TpChr= del.TpChr
			and ins.TpNo= del.TpNo and ins.TpSr= del.TpSr and ins.TpSrNo= del.TpSrNo)
		If @Row > 0
		begin
			Select @ErrMsg='Cannot Change Voucher No.'
			GoTo ErrHndlr
		end
	End
	--DB Perf if @mode in ('C')
	--Begin
	--	select @Row=Count('x') from deleted del where del.TpPrtKey=@wCurrPrtn and
	--		not exists (Select 'x' from inserted ins where ins.TpPrtKey=del.TpPrtKey and ins.TpIdNo=del.TpIdNo
	--		and ins.TpCoCd= del.TpCoCd and ins.TpTc= del.TpTc and ins.TpYy= del.TpYy and ins.TpChr= del.TpChr
	--		and ins.TpNo= del.TpNo and ins.TpSr= del.TpSr and ins.TpSrNo= del.TpSrNo)
	--	If @Row > 0
	--	begin
	--		Select @ErrMsg='Cannot Change Voucher No.'
	--		GoTo ErrHndlr
	--	end
	--End
	
	
	If @mode in ('D')
	Begin
		select @Row= Count('x') from Deleted join Txnd on TdPrtKey=TpPrtKey and TdIdNo=TpTdIdNo
				and TdCoCd= TpCoCd and TdTc= TpTc and TdYy= TpYy and TdChr= TpChr
				and TdNo= TpNo and TdSr= TpSr and TdSrNo= 0
			where TpPrtKey=@wCurrPrtn and not exists (Select 'x' from Bag where BPrtKey=TdPrtKey and BIdno=TdBIdNo 
				and BCoCd= TdCoCd and BYy= TdBYy and BChr= TdBChr and BNo= TdBNo and BLstYyKey= TdYyKey)
		If @Row > 0
		begin
			Select @ErrMsg= 'Cannot Delete Since Not The Last Transaction Of The Bag'
			GoTo ErrHndlr
		end

	End
RETURN
ErrHndlr:
	Begin
		select @errmsg='|'+@errmsg
		raiserror(@errmsg,16,1)
		rollback transaction
	End
END



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





ALTER TRIGGER [dbo].[Txnd_INUPDL] ON [dbo].[Txnd]              
       FOR INSERT,UPDATE,DELETE NOT FOR REPLICATION
AS              
BEGIN              
	Declare @Row Int              
	Declare @Mode Char(1)              
	Declare @ErrMsg Varchar(255)              
	Declare @wTcTyp Char(3)               
	Declare @wTc Char(3)               
	Declare @wEntTyp Char(8)              
	Declare @wHCd Varchar(8)              
	DECLARE @ReplAgt BIT                
	DECLARE @wCurrPrtn VarChar(1)         
	Declare @wTChr Char(1) 
	Declare @AckYN Char(1)  
 	Declare @wSysDbUni Char(1)
	Declare @whOLORYN char(1)
	If @@RowCount=0              
	Return              
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'Txnd' )              
	RETURN              
              
	Select @Mode=Case (Select Count('x') From Inserted)              
	When 0 Then 'D'              
	Else Case (Select Count('x') From Deleted)              
	When 0 Then 'A'              
	Else 'C'              
		End              
	End              
              
	Select @wHCd= 'ZSELF'              
	Select @wCurrPrtn='C'              
      
	Select @wSysDbUni = 'N'
      
	Select @whOLORYN = hOLORYN from head where HCoCd ='ZZZ'
  
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))              

	/*Update 
		Bag 
	set 
		BLabRem='O:'+ISNULL(Del.TdFrBLoc,'')+':N:'+Ins.TdToBLoc+convert(varchar,Ins.TdIdNo)+':'+convert(varchar, getdate(), 120)
	from 
		Inserted Ins
	LEFT OUTER join 
		Deleted Del on Del.TdIdNo=Ins.TdIdNo
	where 
		(ins.TdTc='BV' or ins.TdTc='DT') and 
		(Ins.TdFrBLoc='' and Ins.TdBNo<>0) 
		and Ins.TdSrNo=0
		and Ins.TdBIdNo=Bag.BIdNo
		*/              
	If @Mode in ('A','C')              
	Begin              
	Select @wTcTyp= (PValue) From Inserted Join Param On PTyp='TC'              
		And PMcd= TdTc And PScd=''              
	Select @wTc= TdTc From Inserted Join Param On PTyp='TC'              
		And PMcd= TdTc And PSCd=''              
	Select @wTChr =  (vPValue3) from inserted join vParam on vPCoCd = TdCoCd and vPTyp='CHR'
		And vPMCd = TdTc and vPSCd = TdChr 
	End              
  
               
	If @Mode in ('D')              
	Begin              
	Select @wTcTyp= (PValue) From Deleted Join Param On PTyp='TC'              
		And PMcd=TdTc And PScd=''              
	End              
               
	select @wEntTyp = PValue1 from Param where PTyp='TCTYP'              
     and PMcd=@wTcTyp and PSCd=''              
                
	select @Row= Count('x') from Inserted Ins where Ins.TdSrNo<>0 and
	not exists (select count('x') from Txnd ForChk where ForChk.TdCoCd=Ins.TdCoCd
	and ForChk.TdTc=Ins.TdTc and ForChk.TdYy=Ins.TdYy and ForChk.TdChr = Ins.TdChr and ForChk.TdNo=Ins.TdNo
	and ForChk.TdSr=Ins.TdSr and ForChk.TdSrNo=0)
	If @Row > 0              
	begin              
		Select @ErrMsg= 'Data Issue please contact 9820490905'              
			GoTo ErrHndlr              
	end   

	If @wEntTyp= 'TxnBag'              
	BEGIN        
		if @mode in ('C') 
		begin 
			select @Row= Count('x') from Inserted              where
			not exists (select count('x') from deleted where inserted.TdBIdNo<>deleted.TdBIdNo)
		    If @Row > 0              
			begin              
		      Select @ErrMsg= 'Cannot change Bag number'              
			      GoTo ErrHndlr              
			end   
		end
	--5717
	select @Row= Count('x') from inserted              
    join Bag on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo               
      and TdPrtKey=BPrtKey              
    where BPrtKey=@wCurrPrtn and BLoc= 'PFG' and BRepYN <> 'Y'              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Add/ Change As Bag In Finished Goods Location'              
    GoTo ErrHndlr              
   end  
   --5717
   select @Row= Count('x') from deleted              
    join Bag on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo               
     and TdPRtKey=BPrtKEy              
    where BPrtKey=@wCurrPrtn and BLoc= 'PFG' and BRepYN <> 'Y'              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Change/ Delete As Bag In Finished Goods Location'              
    GoTo ErrHndlr              
   end     
   --5717
   select @Row= Count('x') from inserted where @wTcTyp= 'BV' and TdPRtKey=@wCurrPrtn and              
    (TdFrBLoc= 'PFG' or TdToBLoc= 'PFG')              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'From/ To Bag Location Cannot Be ''PFG'' In Bag Movement'              
    GoTo ErrHndlr              
   end    

   if @mode in ('A','C') 
   begin select @Row= Count('x') from Inserted              
    join Bag on TdBIdNo=BIdNo 
    where BLoc= 'PFG' and BRepYN = 'N'              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot enter, this Bag is In Finished Goods Location'              
    GoTo ErrHndlr              
   end   
   end
   
   if @mode in ('D')           
   begin
   Select @Row= Count('x') from deleted              
    join Bag on TdBIdNo=BIdNo where BLoc= 'PFG' and BRepYN = 'N'              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Delete As Bag In Finished Goods Location'              
    GoTo ErrHndlr              
   end 
   end             

   if @mode in ('A','C')
   begin 
	select @Row= Count('x') from inserted              
    join Bag on TdBIdNo=BIdNo           
    where Not ((TdCoCd= BCoCd and TdPrevYy= BLstYy and TdPrevKey= BLstKey)              
    Or (TdCoCd= BCoCd and TdYy= BLstYy and TdKey= BLstKey))              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Bag in Control of other user'              
    GoTo ErrHndlr              
   end              
   end
	
   if @mode = 'D'
   begin
   select @Row=Count('x') from deleted where 
    Not Exists (Select 'x' from Bag where TdBIdNo=BIdNo  and              
         ((TdYy>= BLstYy and TdKey>= BLstKey)              
           or IsNull(BLstKey, 0)= 0))              
   If @Row > 0              
   Begin              
    Select @ErrMsg= 'Cannot Update Or Delete, Not The Last Txn Of The Bag'              
    GoTo ErrHndlr              
   end  
   end 
              
  If @mode in ('A')              
  Begin
   Select @Row=Count('x') from Inserted               
       Join Bag on TdBIdNo=BIdNo where Tddt<BMaxTxnDt              
   If @Row > 0              
   Begin              
    Select @ErrMsg= 'A Transaction of Later Date Exists For this Bag'              
    GoTo ErrHndlr              
   end     
   select @Row= Count('x') from inserted              
    join Bag on TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo               
      and TdPrtKey=BPrtKey              
    where TdPrtKey=@wCurrPrtn and TdDt<BOpnDt              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Transaction Date Cannot Be Less Than Bag Opening Date'              
    GoTo ErrHndlr              
   end              
         


  End              

 End
   
 
 ----break1				
 --577chk          
 Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPrtn and              
  Not Exists (Select 'x' From Txn Where Txn.TIdNo=Ins.TdTIdNo and Txn.TCoCd=Ins.TdCoCd And Txn.TTc=Ins.TdTc              
   And Txn.TYy=Ins.TdYy And  Txn.TChr=Ins.TdChr And Txn.TNo=Ins.TdNo And Txn.TPrtKey=Ins.TdPrtKey)              
   If @Row> 0              
  begin              
  Select @ErrMsg= 'Voucher No. Does Not Exist'              
  GoTo ErrHndlr              
   end               
 
 --577chk          
 If @wEntTyp= 'TxnBag'              
  Begin              
   /********* Child Recs Should Have The Same Yy-Key & YyKey combination as parent ********/              
   /* Current Partition */              
   select @Row= Count('x') from inserted ins where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo<> 0              
    and exists(Select 'x' from Txnd Td1 where Td1.TdSrNo= 0              
     and (Td1.TdPrtKey=ins.TdPrtKey and Td1.TdTIdNo=ins.TdTIdNo and Td1.TdCoCd= ins.TdCoCd and Td1.TdTc= ins.TdTc              
      and Td1.TdYy= ins.TdYy and Td1.TdChr= ins.TdChr              
      and Td1.TdNo= ins.TdNo and Td1.TdSr= ins.TdSr)              
     and (Td1.TdYy<> ins.TdYy or Td1.TdKey<> ins.TdKey              
      or Td1.TdYyKey<> ins.TdYyKey))              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Yy-Key Combination Of Child Should Be Same As Parent'              
    GoTo ErrHndlr              
   end       
 end       
 --577chk TdToBloc validations
 If @wEntTyp= 'TxnBag'              
  Begin
  select @Row= Count('x') from inserted where @wTcTyp= 'BV' and TdPRtKey=@wCurrPrtn and              
    (TdFrBLoc= 'PFG' or TdToBLoc= 'PFG')              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'From/ To Bag Location Cannot Be ''PFG'' In Bag Movement'              
    GoTo ErrHndlr              
   end  
 
  IF @WTcTyp= 'Bv'              
  BEGIN              
   /* Current Partition */              
   SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn And (TdSrNo=0 AND              
    ((Ins.TdToBLoc <> ISNULL((SELECT Del.TdToBLoc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
     and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
     Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
     Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
    AND              
    NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
      LocTyp= 'P' AND LocCd= Ins.TdToBLoc AND LocValidYN = 'Y'))              
    OR              
    (TdToBLoc = '')))              
   OR              
    (TdSrNo<>0 AND TdToBLoc<> '')              
     IF @Row>0              
   BEGIN              
    SELECT @ErrMsg='Invalid Bag To Loc'              
    GOTO ErrHndlr            
   END              
              
   /* Current Partition */               
   SELECT @row= count('x') FROM inserted Ins WHERE Ins.TdPrtKey=@wCurrPrtn and Ins.TdFrBLoc = Ins.TdToBLoc              
   IF @row>0              
   BEGIN              
    SELECT @errmsg= 'To Bag Loc and From Bag Loc cannot be same'              
    GOTO ErrHndlr              
   END              
  END              
  ELSE              
  BEGIN              
   /* Current Partition */              
   SELECT @Row= COUNT('x') FROM Inserted Ins              
    WHERE TdPrtKey=@wCurrPrtn and TdToBLoc<>''              
     IF @Row>0              
   BEGIN              
    SELECT @ErrMsg='Bag To Loc should be Blank'              
      GOTO ErrHndlr              
   END              
  END          
end
--5717
 If @Mode = 'C'              
 Begin              
  /* Current Partition */                
  select @Row=Count('x') from Deleted Del where TdPrtKey=@wCurrPrtn and              
    Not Exists (Select 'x' from Inserted Ins where Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd=Del.TdCoCd              
     and Ins.TdTc=Del.TdTc and Ins.Tdyy=Del.TdYy and Ins.TdChr= Del.TdChr and Ins.TdNo=Del.TdNo              
     and Ins.TdSr=Del.TdSr and Ins.TdSrNo=Del.TdSrNo and Ins.TdPrtKey=Del.TdPrtKey              
     and Ins.TdByy=Del.TdByy and Ins.TdBChr=Del.TdBChr              
     and Ins.TdBNo=Del.TdBNo and Ins.TdKey= Del.TdKey)              
  If @Row > 0              
  Begin              
   Select @ErrMsg='Cannot Change Company Code/ Voucher No./Bag No/ Key'              
   GoTo ErrHndlr              
  End              
 End  
 --5717
 select @Row= Count(distinct(PValue)) from inserted Join param On              
    PTyp= 'TC' and PMCd= TdTc and PScd= ''              
 If @Row > 1              
 BEGIN              
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'              
  GoTo ErrHndlr              
 END 
 --5717
 select @Row= Count(distinct(PValue)) from deleted Join param on              
    PTyp='TC' and PMcd=TdTc and PSCd=''              
 If @Row > 1              
 BEGIN              
  Select @ErrMsg='Transactions of only One TcType can be Acted on'              
  GoTo ErrHndlr              
 END    
 --5717
 select @wEntTyp = PValue1 from Param where PTyp='TCTYP'              
    and PMcd=@wTcTyp and PSCd=''              
 If @wEntTyp <> 'TxnBag' and @wEntTyp <> 'TxnRm'              
 Begin              
  Select @ErrMsg='Transactions of only ''Txn'' Type permitted'              
  GoTo ErrHndlr              
 end   
 --5717
  Select @Row= Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdSr=0              
 If @Row> 0              
 Begin              
    Select @ErrMsg= 'Sr. Cannot be 0'              
    GoTo ErrHndlr              
 End
 --5717
   Select @Row= Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdKey<= 0              
 If @Row> 0              
 Begin              
  Select @ErrMsg= 'Key Cannot Be Zero'              
  GoTo ErrHndlr              
 End   
 --5717
  Select @Row=Count('x') From Inserted Ins              
  Where Ins.TdPrtKey=@wCurrPrtn and (TdToRmDc=TdFrRmDc And TdFrRmDc<>'')  Or              
   (TdFrRmDc<>'' and TdToRmDc='')              
   If @Row>0              
 Begin              
  Select @ErrMsg='Invalid To Rm Dc'              
  GoTo ErrHndlr              
 End     
 --5717
 Select @Row= count('x') from inserted where TdPrtKey=@wCurrPrtn and (TdFrRmLoc= 'FG' or TdToRmLoc= 'FG')              
  if @Row>0               
  Begin               
   Select @ErrMsg= 'Cannot Make Any Transaction For ''FG'' Location Type'              
   GoTo ErrHndlr              
  End
  --5717
  Select @Row= count('x') from inserted where TdPrtKey=@wCurrPrtn       
   and (TdFrRmLoc= 'XD' or TdToRmLoc= 'XD'       
      Or TdFrRmLocTyp In ('XD', 'MG', 'Y') or TdToRmLocTyp In ('XD', 'MG', 'Y'))      
  if @Row>0               
  Begin               
   Select @ErrMsg= 'Cannot Make Any Transaction For ''XD'',''MG'',''Y'' Location Type'              
   GoTo ErrHndlr              
  End   
  --5717
  select @Row= Count('x') from Inserted Ins where Ins.TdPrtKey=@wCurrPRtn and Ins.TdRtByQW <>'' and @wTc<>'PR'     
and @wTC<>'MV' and @wTc<>'CNV' and @wTc<>'SA' 
   If @Row> 0              
   begin              
    Select @ErrMsg= 'Rm Qw Option Should be Blank'              
    GoTo ErrHndlr              
   end
--5717
If @Mode= 'A'              
 Begin    
 select @Row= Count('x') from inserted ins where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo= 0              
   and exists(Select 'x' from Txnd Td1 where Td1.TdSrNo= 0              
    and Td1.TdCoCd= ins.TdCoCd and Td1.TdYy= ins.TdYy              
    and Td1.TdKey= ins.TdKey and              
    (Td1.TdCoCd<> ins.TdCoCd or Td1.TdTc<> ins.TdTc or              
     Td1.TdYy<> ins.TdYy or Td1.TdChr<> ins.TdChr or              
     Td1.TdNo<> ins.TdNo or Td1.TdSr<> ins.TdSr))              
    If @Row> 0              
  begin              
   Select @ErrMsg= 'Yy-Key Combination Already Exists For Another Voucher Sr.'              
   GoTo ErrHndlr              
  end 
  --5717
   If @wEntTyp= 'TxnBag'              
  Begin 
  select @Row= Count('x') from inserted ins where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo= 0              
    and exists(Select 'x' from Txnd Td1 where Td1.TdSrNo= 0              
     and Td1.TdBIdNo=ins.TdBIdNo and Td1.TdCoCd= ins.TdCoCd and Td1.TdBYy= ins.TdBYy              
     and Td1.TdBChr= ins.TdBChr  and Td1.TdBNo= ins.TdBNo              
     and (Td1.TdYy> ins.TdYy              
       Or              
      (Td1.TdYy= ins.TdYy and Td1.TdKey> ins.TdKey)))              
     If @Row> 0              
   Begin              
    Select @ErrMsg= 'A Later Txn Of The Bag Exists, Try Entering This Rec Again'              
    GoTo ErrHndlr              
   End    
   end
 end          
 --5717
 IF @wEntTyp='TxnBag'              
 Begin
 Select @Row=Count('x') From Inserted Ins              
     Where TdPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Param where PTyp='DC'              
    and PMCd = Ins.TdFrRmDc and PScd='')              
    and TdSrNo<>0)              
   OR              
    (TdSrNo=0 and TdFrRmDc<>'')              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid From Rm Ir/Rm Ir Should Not Be Entered'              
     GoTo ErrHndlr              
  END
 
 --5717
 SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and (TdSrNo<>0 AND              
   NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
      LocTyp<> 'P' and LocTyp <>'R' AND LocCd = Ins.TdToRmLoc AND LocValidYN='Y' And              
       (@wTcTyp NOT IN('Rec','Cnv','Bs', 'BM') OR LocTyp= 'Xc')))              
   OR              
    (TdSrNo=0 AND TdToRmLoc<>'')              
    IF @Row> 0              
  BEGIN              
   SELECT @ErrMsg= 'Invalid To Rm Loc/ To Rm Loc Should Be Blank'              
   GOTO ErrHndlr              
  END    
  --5717
   SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and (TdSrNo<>0 AND              
   ((              
   (Ins.TdToRmLoc <> ISNULL((SELECT Del.TdToRmLoc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
    and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   Or              
   Ins.TdToRmDc <> ISNULL((SELECT Del.TdToRmDc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
    and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   )              
   AND TdToRmDc = 'D'              
   AND NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
      LocTyp<> 'P' and LocTyp<> 'R' AND LocCd = Ins.TdToRmLoc AND LocValidYN = 'Y' AND              
       (@wTcTyp NOT IN('Rec','Cnv','Bs','BM') OR LocTyp= 'Xc'))) OR (TdToRmLoc = '')))              
   OR              
    (TdSrNo=0 AND TdToRmLoc<>'')              
  /* **** Zubin 211 (Or Condition for 'TdToRmLoc' was missing) **** */              
    IF @Row> 0              
  BEGIN              
   SELECT @ErrMsg= 'Invalid To Rm Loc/ To Rm Loc Should Be Blank'              
   GOTO ErrHndlr              
  END        
  --5717
    Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and (@wTcTyp='MLT' and TdSrNo=0 and TdRjDc<>'C') Or              
    (@wTcTyp='REJ' and TdSrNo=0 and TdRjDc=TdBDc)              
   Or              
    ((TdSrNo<>0 or @wTctyp Not In('MLT','REJ')) and TdRjDc<>'')              
    If @Row>0             
  BEGIN              
   Select @ErrMsg='Invalid Rej Ir/Rej Ir Should Not Be Entered'              
   GoTo ErrHndlr              
  END 
  --5717
    Select @Row=Count('x') From Inserted Ins              
     Where TdPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Param where PTyp='YN'              
    and PMCd = Ins.TdAck and PScd='')              
   and  @wTcTyp='Bv')                 
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Bag Movement Ack/Ack Should Not Be Entered'              
     GoTo ErrHndlr              
  END 
  --5717
    Select @Row=Count('x') From Inserted Ins              
  Where TdPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Param where PTyp='YN'              
       and PMCd = Ins.TdPrdYN and PScd='')              
   and  @wTcTyp In('Dt') and TdSrNo=0)              
   OR              
 (@wTcTyp Not In('Dt') and TdPrdYN<> '')              
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Prd YN/Prd YN Should Not Be Entered'              
     GoTo ErrHndlr              
  END 
  --5717
    SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and              
   Ins.TdByLoc <> ISNULL((SELECT Del.TdByLoc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
     and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   AND              
   NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
     LocTyp= 'P' AND LocCd = Ins.TdByLoc AND LocValidYN = 'Y')              
   AND TdByLoc <> ''              
    IF @Row> 0              
  BEGIN              
   SELECT @ErrMsg= 'Invalid By Loc'              
     GOTO ErrHndlr              
  END  
  --5717
    SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and               
   Ins.TdByWrk <> ISNULL((SELECT Del.TdByWrk FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
    and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   AND              
     NOT EXISTS (SELECT 'x' FROM vParam WHERE vPCoCd = TdCoCd              
     AND vPTyp='WORK' AND vPMCd = Ins.TdByWrk AND vPValidYN = 'Y')              
   AND TdByWrk<>''              
      IF @Row>0              
  BEGIN              
   SELECT @ErrMsg='Invalid By Work'              
     GOTO ErrHndlr              
  END 
  --5717
    Select @Row=Count('x') From Inserted Ins              
     Where TdPrtKey=@wCurrPrtn and (TdPrdYN= 'Y' and (TdByLoc= '' or TdByWrk= '')) or              
               (IsNull( (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
     LocTyp in ('L', 'B', 'M')  and LocCd = TdToRmLoc), '')<> '' and (TdByLoc= '' or TdByWrk= '')              
    )              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='By Loc/ Wrk Should Be Entered'              
     GoTo ErrHndlr              
  END   
  --5717
    Select @Row= Count('x') from inserted Ins       
  Join Txn On TPrtKey=Ins.TdPRtKey and TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where TdPrtKey=@wCurrPrtn and              
   (@wTcTyp not in ('BS', 'BM') or TdSrNo<> 0 or (((TMultIIr<>'I' and TdBDc<> 'D') or (TMultIIr='I' and TdBDc<> 'C')) and @wTcTyp<> 'BM') ) and               
   (TdRefYy<> '' or TdRefKey<> 0)                 
      If @Row > 0               
       Begin              
   Select @ErrMsg='Ref Yy && Ref Key should be blank'              
     GoTo ErrHndlr              
    end  
	--5717
	 Select @Row=Count('x') from inserted Ins where TdPrtKey=@wCurrPrtn and              
   (Not Exists (Select 'x' from Txnd  where Txnd.TdPrtKey=Ins.TdPrtKey and Txnd.TdBIdNo=Ins.TdBIdNo and Txnd.TdCoCd=Ins.TdCoCd and              
    Txnd.TdBYy=Ins.TdBYy and Txnd.TdBChr=Ins.TdBChr and Txnd.TdBNo=Ins.TdBNo              
    and Txnd.TdYy=Ins.TdPrevYy and Txnd.TdKey=Ins.TdPrevKey)              
   and TdSrno=0)              
   and              
   (TdPrevYy<>'' or TdPrevKey<>0)              
    If @Row > 0              
     Begin              
   Select @ErrMsg='Invalid Prev YY/Key'              
   GoTo ErrHndlr              
  end
  --5717
    if @mode in ('A','C')              
  begin                 
   Select @Row=Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and              
    Exists (Select 'x' From Bag Where BPrtKey=Ins.TdPrtKey and BIdNo=Ins.TdBIdNo and BCoCd= Ins.TdCoCd and              
     BYy= Ins.TdByy and BChr= Ins.TdBChr and BNo= Ins.TdBNo and Cast(BQty as Decimal(8,2))< 0)              
     If @Row>0              
   begin              
    Select @ErrMsg='Cannot accept Txn, Bag Qty going -ve'              
    GoTo ErrHndlr              
   end              
  end              
  --5717            
  if @Mode='D'              
  Begin                 
   Select @Row=Count('x') From Deleted Del Where Del.TdPRtKey=@wCurrPrtn and               
    Exists (Select 'x' From Bag Where BPrtKey=Del.TdPrtKey and BIdNo=Del.TdBIdNo and BCoCd= Del.TdCoCd              
         and BYy= Del.TdByy and BChr= Del.TdBChr              
         and BNo = Del.TdBNo and Cast(BQty as Decimal(8,2))< 0)              
     If @Row>0              
     begin              
    Select @ErrMsg='Cannot accept Txn, Bag Qty going -ve'              
    GoTo ErrHndlr              
   end              
  End
  --5717
    Select @Row=Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPrtn and Ins.TdRmCd<>'' and              
   IsNull((Select sum(Td.TdRmQty* (              
        Case when Td.TdFrRmLoc= 'W'  and Td.TdFrRmDc= 'D' then 1              
     when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
     when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
     when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
            Else 0 end))              
     from Txnd Td where Td.TdPrtKey=Ins.TdPrtKey and Td.TdBIdNo=Ins.TdBIdNo               
     and Td.TdCoCd= Ins.TdCoCd and Td.TdBYy= Ins.TdBYy              
     and Td.TdBChr= Ins.TdBChr and Td.TdBNo= Ins.TdBNo              
     and Td.TdRmCd= Ins.TdRmCd and Td.TdLotNo= Ins.TdLotNo              
     and Str(Td.TdRmSz, 16, 4)= Str(Ins.TdRmSz, 16, 4)              
     and Str(Td.TdRmSz2, 16, 4)= Str(Ins.TdRmSz2, 16, 4)              
     and Str(Td.TdRmSz3, 16, 4)= Str(Ins.TdRmSz3, 16, 4)              
     and Str(Td.TdRmStkRt, 16, 2)= Str(Ins.TdRmStkRt, 16, 2)), 0)< 0              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Rm Qty Is Less Than Zero In Bag'              
   GoTo ErrHndlr              
  END    
  --5717
  If @Mode in ('C')              
  Begin              
   Select @Row= Count('x') From Deleted Del Where Del.TdPrtKey=@wCurrPrtn and Del.TdRmCd<> '' and               
    exists(Select 'x' from Inserted Ins where               
  Ins.TdPrtKey=Del.TdPRtKey and Ins.TdBIdNo=Del.TdBIdNo              
     and Ins.TdCoCd= Del.TdCoCd and Ins.TdBYy= Del.TdBYy              
     and Ins.TdBChr= Del.TdBChr and Ins.TdBNo= Del.TdBNo              
     and (Ins.TdRmCd<> Del.TdRmCd or Ins.TdLotNo<> Del.TdLotNo              
     or Str(Ins.TdRmSz, 16, 3)<> Str(Del.TdRmSz, 16, 3)              
     or Str(Ins.TdRmSz2, 16, 3)<> Str(Del.TdRmSz2, 16, 3)              
     or Str(Ins.TdRmSz3, 16, 3)<> Str(Del.TdRmSz3, 16, 3)              
     or Str(Ins.TdRmStkRt, 16, 2)<> Str(Del.TdRmStkRt, 16, 2))) and              
     Round(IsNull((Select sum(Td. TdRmQty* (              
     Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
     Else 0 end))              
    from Txnd Td where Td.TdPrtKey=Del.TdPRtKey and Td.TdBIdNo=Del.TdBIdNo              
    and Td.TdCoCd= Del.TdCoCd and Td.TdBYy= Del.TdBYy              
    and Td.TdBChr= Del.TdBChr and Td.TdBNo= Del.TdBNo              
    and Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
    and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
    and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
	and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
    and Str(Td.TdRmStkRt, 16, 2)= Str(Del.TdRmStkRt, 16, 2)), 0), 3)< 0               
  If @Row>0              
  Begin              
   Select @ErrMsg='Rm Qty Is Less Than Zero In Bag'              
   GoTo ErrHndlr              
  End            
  End

  --5717
  if @Mode='D'              
  Begin              
   Select @Row=Count('x') From Deleted Del Where Del.TdPrtKey=@wCurrPrtn and Del.TdRmCd<>'' and              
    IsNull((Select sum(Td.TdRmQty* (              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end))            
      from Txnd Td where Td.TdPrtKey=Del.TdPrtKey and Td.TdBIdNo=Del.TdBIdNo              
      and Td.TdCoCd= Del.TdCoCd and Td.TdBYy= Del.TdBYy              
      and Td.TdBChr= Del.TdBChr and Td.TdBNo= Del.TdBNo              
      and Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
      and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
	  and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
	  and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
      and Str(Td.TdRmStkRt, 16, 2)= Str(Del.TdRmStkRt, 16, 2)), 0)< 0              
     If @Row>0              
   begin              
    Select @ErrMsg='Rm Qty Is Less Than Zero In Bag'              
    GoTo ErrHndlr              
   end           
   End
   --5717
    Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPrtn and Ins.TdRmCd<> '' and              
   Round(IsNull((Select sum(Td.TdRmWt* (              
        Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
     when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
     when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
     when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
            Else 0 end))              
     from Txnd Td where Td.TdPrtKey=Ins.TdPrtKey and Td.TdBIdNo=Ins.TdBIdNo              
     and Td.TdCoCd= Ins.TdCoCd and Td.TdBYy= Ins.TdBYy              
     and Td.TdBChr= Ins.TdBChr and Td.TdBNo= Ins.TdBNo              
     and Td.TdRmCd= Ins.TdRmCd and Td.TdLotNo= Ins.TdLotNo              
     and Str(Td.TdRmSz, 16, 4)= Str(Ins.TdRmSz, 16, 4)              
     and Str(Td.TdRmSz2, 16, 4)= Str(Ins.TdRmSz2, 16, 4)              
	 and Str(Td.TdRmSz3, 16, 4)= Str(Ins.TdRmSz3, 16, 4)              
     and Str(Td.TdRmStkRt, 16, 2)= Str(Ins.TdRmStkRt, 16, 2)), 0), 3)< 0              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Rm Wt Is Less Than Zero In Bag'              
   GoTo ErrHndlr              
  END    
  --5717
    If @Mode in ('C')              
  Begin              
   Select @Row= Count('x') From Deleted Del Where Del.TdPrtKey=@wCurrPrtn and Del.TdRmCd<> '' and               
    exists(Select 'x' from Inserted Ins where            
     Ins.TdPrtKey=Del.TdPRtKey and Ins.TdBIdNo=Del.TdBIdNo              
     and Ins.TdCoCd= Del.TdCoCd and Ins.TdBYy= Del.TdBYy              
     and Ins.TdBChr= Del.TdBChr and Ins.TdBNo= Del.TdBNo              
     and (Ins.TdRmCd<> Del.TdRmCd or Ins.TdLotNo<> Del.TdLotNo              
     or Str(Ins.TdRmSz, 16, 3)<> Str(Del.TdRmSz, 16, 3)              
     or Str(Ins.TdRmSz2, 16, 3)<> Str(Del.TdRmSz2, 16, 3)              
     or Str(Ins.TdRmSz3, 16, 3)<> Str(Del.TdRmSz3, 16, 3)              
     or Str(Ins.TdRmStkRt, 16, 2)<> Str(Del.TdRmStkRt, 16, 2))) and              
     Round(IsNull((Select sum(Td.TdRmWt* (              
     Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
     Else 0 end))              
    from Txnd Td where Td.TdPrtKey=Del.TdPRtKey and Td.TdBIdNo=Del.TdBIdNo              
    and Td.TdCoCd= Del.TdCoCd and Td.TdBYy= Del.TdBYy              
    and Td.TdBChr= Del.TdBChr and Td.TdBNo= Del.TdBNo              
    and Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
    and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
    and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
    and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
    and Str(Td.TdRmStkRt, 16, 2)= Str(Del.TdRmStkRt, 16, 2)), 0), 3)< 0               
  If @Row>0              
  Begin              
   Select @ErrMsg='Rm Wt Is Less Than Zero In Bag'              
   GoTo ErrHndlr              
  End            
  End
  --5717
    if @Mode='D'              
  Begin              
   Select @Row= Count('x') From Deleted Del Where Del.TdPrtKey=@wCurrPrtn and Del.TdRmCd<> '' and              
    Round(IsNull((Select sum(Td.TdRmWt* (              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end))              
      from Txnd Td where Td.TdPrtKey=Del.TdPRtKey and Td.TdBIdNo=Del.TdBIdNo              
      and Td.TdCoCd= Del.TdCoCd and Td.TdBYy= Del.TdBYy              
      and Td.TdBChr= Del.TdBChr and Td.TdBNo= Del.TdBNo              
      and Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
      and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
      and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
      and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
      and Str(Td.TdRmStkRt, 16, 2)= Str(Del.TdRmStkRt, 16, 2)), 0), 3)< 0              
     If @Row>0              
   begin              
    Select @ErrMsg='Rm Wt Is Less Than Zero In Bag'              
    GoTo ErrHndlr              
   end        
   End
   --5717
     Select @Row=Count('x') From Inserted Ins              
   Where TdPRtKey=@wCurrPrtn               
   and Not Exists (Select 'x' From Bag Where BPrtKey=Ins.TdPrtKey and BIdNo=Ins.TdBIdNo              
       and BCoCd= Ins.TdCoCd              
       and BYy= Ins.TdByy and BChr=Ins.TdBChr              
       and BNo = Ins.TdBNo and BRjQty>= 0)              
    If @Row>0              
  BEGIN              
   Select @ErrMsg= 'Cannot accept Bag Rej Qty going -ve'              
   GoTo ErrHndlr              
  END     
 end 
  --5717 
  If @wEntTyp= 'TxnBag'              
  BEGIN    
    select @Row= Count('x') from inserted Ins       
  Join Txn On TPrtKey=Ins.TdPRtKey and TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where @wTcTyp= 'BS' and TdPRtKEy=@wCurrPrtn and TMultiIr<>'I'  and            
   (TdRefYy<> '' or TdRefKey<> 0) and               
   Not Exists (Select 'x' from Txnd Td where Td.TdPrtKey=Ins.TdPrtKey and Td.TdTIdNo=Ins.TdTIdNo               
   and Td.TdCoCd= Ins.TdCoCd and Td.TdTc= Ins.TdTc and Td.Tdyy= Ins.TdYy               
   and Td.TdChr= Ins.TdChr and Td.TdNo= Ins.TdNo               
   and Td.TdYy= Ins.TdRefYy and Td.TdKey= Ins.TdRefKey)      
  /*** (Jen 2.13) ***/              
  If @Row > 0              
  Begin              
   Select @ErrMsg= 'RefYy, RefKey is not in the Voucher'              
   GoTo ErrHndlr              
  end   
  --5717
    select @Row=Count('x') from inserted Ins where Ins.TdPrtKey=@wCurrPrtn and Ins.TdSrNo<>0 and              
   Not Exists (Select 'x' from Txnd Td               
   where Td.TdPRtkey=Ins.TdPrtKey and Td.TdTIdNo=Ins.TdTIdNo               
   and Td.TdCoCd=Ins.TdCoCd and Td.TdTc=Ins.TdTc and Td.Tdyy=Ins.Tdyy              
   and Td.TdChr= Ins.TdChr and Td.TdNo=Ins.TdNo              
   and Td.TdSr=Ins.TdSr and Td.TdSrNo=0              
   and Td.TdBIdNo=Ins.TdBIdNo and Td.TdBYy= Ins.TdBYy and Td.TdBChr= Ins.TdBChr              
   and Td.TdBNo= Ins.TdBNo)              
  If @Row > 0              
  Begin              
   Select @ErrMsg='TdSrNo=0 record does not exist'              
   GoTo ErrHndlr              
  end 
  --5717
    If @wTcTyp= 'BS'              
  Begin   
     Select @Row= Count('x') From Inserted Ins              
    Join Param P1 On P1.PTyp= 'BCHR' and P1.PMCd= ins.TdBChr and P1.PSCd= ''              
    where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo= 0 and          
     Exists (Select 'x' from Txnd Td              
      Join Param P2 On P2.PTyp= 'BCHR' and P2.PMCd= Td.TdBChr and P2.PSCd= ''              
      where Td.TdPRtKey=Ins.TdPRtKey and Td.TdTIdNo=Ins.TdTIdNo and Td.TdCoCd= Ins.TdCoCd and              
      Td.TdTc= Ins.TdTc and Td.TdYy= Ins.TdYy and Td.TdChr= Ins.TdChr and              
      Td.TdNo= Ins.TdNo and Td.TdSrNo= 0 and P2.PValue<> P1.PValue )              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Bag Split Can Take Place Between Same Bag Type'              
    GoTo ErrHndlr              
   end       
   --5717
      select @Row= Count('x') from inserted Ins      
  Join Txn On TPrtKey=Ins.TdPRtKey and TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and ((TMultiIr<>'I' and TdBDc= 'D') /*or (TMultiIr='I' and TdBDc= 'C')*/)      
    and (TdRefYy= '' or TdRefKey<= 0)              
       If @Row > 0              
        begin              
    Select @ErrMsg= 'RefYy / RefKey cannot be Blank'              
      GoTo ErrHndlr              
     end   
	   
	end
  END
              
              
 If @Mode ='D' and @wEntTyp='TxnBag'              
  Begin              
   Select @Row=Count('x') from Deleted Del where TdSrNo=0 and              
      Exists (Select 'x' from Txnd Td Where Td.TdTIdNo=Del.TdTIdNo               
    And Td.TdCoCd= Del.TdCoCd And Td.TdTc= Del.TdTc and Td.Tdyy= Del.Tdyy              
    And Td.TdChr= Del.TdChr and Td.TdNo= Del.TdNo              
    And Td.TdSr= Del.TdSr and Td.TdSrNo<> 0)              
   If @Row > 0              
   Begin              
    Select @ErrMsg='Cannot Delete SrNo=0 Record, As Childs Exist'              
    GoTo ErrHndlr              
   End              
              
   Select @Row=Count('x') from Deleted Del where  TdSrNo=0 and              
      Exists (Select 'x' from TxndPrd where TpTdIdNo=TdIdNo and TpCoCd= TdCoCd and TpTc= TdTc              
     and TpYy= TdYy and TpChr= TdChr and TpNo= TdNo and TpSr= TdSr)              
   If @Row > 0              
   Begin              
    Select @ErrMsg='Cannot Delete SrNo=0 Record, As Prd Points Record Exists'              
    GoTo ErrHndlr              
   End           
  End              
  --5717
  if @whOLORYN= 'Y'
 begin
  select @Row= Count('x') from inserted ins where TdRmCtg in ('D','C','X') and TdLotNo =''              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Lot No. should Be Entered For Diamonds,Color Stones,Accessories'              
   GoTo ErrHndlr              
  End
  --5717
  select @Row= Count('x') from inserted ins where TdRmCtg in ('D','C','X') and @wEntTyp<> 'TxnBag' and TdLotNo <>'' and TdPurRt=0             
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Rate cannot be zero, Rate should Be Entered For Diamonds,Color Stones,Accessories'              
   GoTo ErrHndlr              
  End
  --5717
  Select @Row = Count('x') from inserted ins where TdRmCtg not in ('D','C','X') and TdLotNo <> ''              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Lot No. should not be entered for other than Diamonds,Color Stones,Accessories'              
   GoTo ErrHndlr              
  End
  --5717
  Select @Row=Count('x') from Deleted Del where TdLotMstRecYN='Y' and @mode='D' and                
      Exists (Select 'x' from Txnd where TdLotNo = Del.TdLotNo and TdIdNo <> Del.TdIdNo)              
   If @Row > 0              
   Begin              
    Select @ErrMsg='Cannot Delete Lot Master record as child transactions exist for this Lot No'              
    GoTo ErrHndlr              
   End  
   --5717
   If @Mode = 'C'              
   Begin                 
    select @Row=Count('x') from Deleted Del where TdLotMstRecYN='Y' and exists (Select 'x' from Txnd where TdLotNo = Del.TdLotNo and TdIdNo <> Del.TdIdNo)             
        If @Row > 0              
   Begin              
    Select @ErrMsg='Cannot Change RmCd/LotNo/RtByQW/Rt for a Lot Master record as child transactions exist for this Lot No'              
    GoTo ErrHndlr              
   End              
  End    
End
 --5717
  If @wEntTyp in ('TxnRm', 'TxnBag')              
 Begin                
  select @Row= Count('x') from inserted ins  where @whOLORYN ='N' and TdPRtKey=@wCurrPrtn and TdLotNo <> '' and              
   Not Exists (select 'x' from RmMst where RmZ= 'N' and RmPrtKey=TdPrtKey and RmCd= TdRmCd and RmCtg IN ('D', 'C', 'X','G','P','S','L','M'))              
    If @Row> 0              
  begin              
   Select @ErrMsg= 'Lot No. should Be Entered For Diamonds ,Color Stones,Accessories, Gold, Palladium, Silver, Platinum and Metal Only'              
   GoTo ErrHndlr              
  end  
  --5717
  Select @Row= Count('x') From Inserted Ins Where @wholoryn='N' and TdPrtKey=@wCurrPRtn and TdLotNo<> '' and              
   Not Exists (Select 'x' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and HLotNoYN= 'Y')              
    If @Row> 0              
  begin              
   Select @ErrMsg= 'Lot No. should Be Blank As ''Implement Lot No. '' Option Is ''N'' In Head Master (Tab 5)'              
     GoTo ErrHndlr              
  end              
 End      
    

  If @wTcTyp In('DT', 'BS')              
  Begin              
	if @Mode='A'
	begin
		Select @Row= Count('x') from Inserted Ins              
			Join RmMst On RmCd= Ins.TdRmCd and RmPrtKey=Ins.TdPrtKey              
			Join Param ON PTyp= 'BCHR' and PMCd= Ins.TdBChr and PSCd= '' and PValue in ('F')               
			where Ins.TdPrtKey=@wCurrPRtn and Ins.TdSrNo> 0 and              
				NOT (              
				 (',' + PVALUE1 + ',') LIKE (CASE WHEN PVALUE1='' THEN '%' ELSE '%,'+RmCtg+',%' END) and              
				 (',' + PVALUE2 + ',') LIKE (CASE WHEN PVALUE2='' THEN '%' ELSE '%,'+RmSCtg+',%' END) and
				 ((Pdesc225='' and PValue3='') or (exists (Select 'a' from Bag 
																		join OrdRm on Bag.BOdIdNo = OrdRm.OrOdIdNo 
																		where Bag.BCoCd=Ins.TdCoCd and Bag.BYy = Ins.TdByy and Bag.BChr = Ins.TdBChr and Bag.BNo = Ins.TdBNo and TdRmCd =OrRmCd and
																		(PDesc225='' or PDesc225=OrSubShp)
																		and (PValue3='' or (PValue3='WS' and OrWsQty>0) or (PValue3='GS' and OrHsQty>0))
																		 					))))
				
	   If @Row> 0              
		begin              
			Select @ErrMsg= 'Cannot Issue Raw Material Of This Ctg/Sub-Category To This Flute Bag'              
			GoTo ErrHndlr              
		end              
	end
 End              
 
   If @wTcTyp In('DT')              
  Begin              
   --5717
   select @Row= Count('x') from Inserted Ins              
    where Ins.TdPrtKey=@wCurrPrtn and (Ins.TdFBRmQty<> '' or Ins.TdFBRmWt<> '' ) and              
     Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= Ins.TdBChr and              
      PSCd= '' and PValue in ('C', 'P'))              
   If @Row> 0              
   begin              
    Select @ErrMsg= 'Cannot Enter RM Quantity/Weight Reference For Primary/ Component Bags'              
    GoTo ErrHndlr              
   end
   /* Zubin 212 */              
   /* Current Partition */              
	if @wTChr = 'Y'
	begin
     select @row= count('x') from inserted where TdPrtKey=@wCurrPRtn and TdSrNo <> 0 and              
     TdFrRmLoc<> (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'XR') AND              
     TdToRmLoc<> (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'XR')              
     if @row> 0              
		begin              
			select @errmsg= 'Either RmLocFr Or RmLocTo Should Belong To ''XR'' Loc Type'              
			goto ErrHndlr              
		end              


     select @row= count('x') from inserted where TdPrtKey=@wCurrPRtn and TdSrNo <> 0 and              
     TdFrRmLoc= (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'W') Or              
     TdToRmLoc= (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'W')              
     if @row> 0              
		begin              
			select @errmsg= 'RmLocFr Or RmLocTo cannot Belong To ''W'' Loc Type'              
			goto ErrHndlr              
		end              
	END
	else
	begin
	     select @row= count('x') from inserted where TdPrtKey=@wCurrPRtn and TdSrNo <> 0 and              
			TdFrRmLoc<> (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'W') AND              
			TdToRmLoc<> (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'W')              
			if @row> 0              
			begin              
				select @errmsg= 'Either RmLocFr Or RmLocTo Should Belong To ''W'' Loc Type'              
				goto ErrHndlr              
			end              
	
	 select @row= count('x') from inserted where TdPrtKey=@wCurrPRtn and TdSrNo <> 0 and              
     TdFrRmLoc= (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'XR') Or              
     TdToRmLoc= (select LocCd from Loc where LocCoCd= TdCoCd and LocTyp= 'XR')              
     if @row> 0              
		begin              
			select @errmsg= 'RmLocFr Or RmLocTo cannot Belong To ''XR'' Loc Type'              
			goto ErrHndlr              
		end              
	ENd
   /* Zubin 212 */              
  END
 --5717
   If @wTcTyp Not In('DT')              
  Begin              
   /* Current Partition */               
   select @Row= Count('x') from Inserted where TdPrtKey=@wCurrPrtn and TdFBRmQty<> '' or TdFBRmWt<> ''              
   If @Row> 0              
   begin              
    Select @ErrMsg= 'Cannot Enter RM Quantity/Weight Reference For This TcTyp'              
    GoTo ErrHndlr              
   end              
  End   
  --5717
  if @mode = 'C'              
  Begin
  If @wTcTyp= 'DT'              
   begin              
    /* Current Partition */              
    select @Row= Count('x') from Inserted Ins where TdPRtKey=@wCurrPrtn and   
     Exists (Select 'x' from deleted del where ins.TdPrtKey=del.TdPrtKey and ins.TdIdNo=del.TdIdNo               
     and ins.TdCoCd= del.TdCoCd and ins.TdTc= del.TdTc and ins.TdYy= del.TdYy              
     and ins.TdChr= del.TdChr and ins.TdNo= del.TdNo              
     and ins.TdSr = del.TdSr and ins.TdSrNo = del.TdSrNo and              
      (ins.TdByWrk<> del.TdByWrk or ins.TdByLoc<> del.TdByLoc))              
     and              
     Exists (Select 'x' from TxndPrd where TpPrtKey=ins.TdPrtKey               
      and TpTdIdNo=ins.TdIdNo and TpCoCd= ins.TdCoCd              
      and TpTc= ins.TdTc and TpYy= ins.TdYy and TpChr= ins.TdChr              
      and TpNo= ins.TdNo and TpSr= ins.TdSr)              
    If @Row > 0              
    begin              
     Select @ErrMsg= 'Cannot Change Prd as Production Points Exist'              
     GoTo ErrHndlr              
    end              
   end         
   select @Row= Count('x') from inserted ins where ins.TdPrtKey=@wCurrPrtn and ins.TdTc = 'BV' and              
   not exists(Select 'X' from deleted del where del.TdPrtKey=ins.TdPrtKey and del.TdIdNo=ins.TdIdNo               
   and del.Tdcocd= ins.TdCoCd and del.TdTc= ins.TdTc and del.TdYy= ins.TdYy and del.TdChr= ins.TdChr              
   and del.TdNo= ins.TdNo and del.TdSr= ins.TdSr and del.TdBYy= ins.TdBYy              
   and del.TdBChr= ins.TdBChr and del.TdBNo= ins.TdBNo and del.TdBQty= ins.tdBQty              
   and del.TdBDc= ins.TdBDc and del.TdBGrWt= ins.TdBGrWt and del.TdFrBLoc= ins.TdFrBLoc              
   and del.TdToBLoc= ins.TdToBLoc and del.TdFrRmLoc= ins.TdFrRmLoc and del.TdFrRmDc= ins.TdFrRmDc              
   and del.TdRmCd= ins.TdRmCd and del.TdLotNo= ins.TdLotNo and del.TdRmSz= ins.TdRmSz              and
   del.TdRmSz2= ins.TdRmSz2  and del.TdRmSz3= ins.TdRmSz3                          
   and del.TdRmStkRt= ins.TdRmStkRt              
   and del.TdRmQty= ins.TdRmQty and del.TdRmWt= ins.TdRmWt and del.TdToRmLoc= ins.TdToRmLoc              
   and del.TdToRmDc= ins.TdToRmDc and del.TdWrk= ins.TdWrk and del.TdPrdYN= ins.TdPrdYN              
   and del.TdByLoc= ins.TdByLoc and del.TdByWrk= ins.TdByWrk and del.TdDustWt= ins.TdDustWt and              
   del.TdKey= ins.TdKey and del.TdRefYy= ins.TdRefYy and del.TdRefKey= ins.TdRefKey and              
   del.TdRjCd= ins.TdRjCd and del.TdRjQty= ins.TdRjQty and del.TdRjDc= ins.TdRjDc and              
   del.TdPtQty= ins.TdPtQty and del.TdRjPtQty= ins.TdRjPtQty and del.TdDesc= ins.TdDesc              
   and del.TdPurRt= ins.TdPurRt and del.TdPurAmt= ins.TdPurAmt and del.TdPrevYy= ins.TdPrevYy              
   and del.TdPrevKey= ins.TdPrevKey and del.TdYyKey= ins.TdYyKey and del.TdPcPerCt= ins.TdPcPerCt )              
   If @Row > 0              
   Begin              
    Select @ErrMsg= 'Cannot Update any as Bag is Acknowledged'              
    GoTo ErrHndlr              
   end              
  End 
  
  
             
 If @wEntTyp= 'TxnBag'              
 Begin              
          
              
  /**** Cannot Update or Delete if Fg Txn later to the Txn Entry of the Bag exists *******/              
  /* Current Partition */               
  if @mode='D'
  begin 
   Select @Row=Count('x') from deleted where 
   Exists (Select 'x' from Fgd where FdBIdNo=TdBIdNo and FdCoCd= TdCoCd and FdBYy= TdBYy              and FdBChr= TdBChr and FdBNo= TdBNo and FdBYyKey>= TdYyKey)              
  If @Row > 0              
  begin              
   Select @ErrMsg= 'Cannot Update Or Delete, A Later Fg Txn Exists For The Bag'              
   GoTo ErrHndlr              
  end
  end
  
  if @mode in ('A','C')
  begin 
	Select @Row=Count('x') from inserted where TdSrNo<> 0 and              
    Exists (Select 'x' from Fgd where FdBIdNo=TdBIdNo and FdCoCd= TdCoCd and FdBYy= TdBYy              
   and FdBChr= TdBChr and FdBNo= TdBNo and FdBYyKey>= TdYyKey)              
  If @Row > 0              
  begin              
   Select @ErrMsg= 'Cannot Add Or Update, A Later Fg Txn Exists For The Bag'              
   GoTo ErrHndlr              
  end
  
  Select @Row=Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and              
   Not Exists (Select 'x' From Loc where LocCoCd= TdCoCd and              
   LocTyp= 'P' and LocCd = Ins.TdFrBLoc)              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Bag Fr Loc, contact 9820490905 right away'              
     GoTo ErrHndlr              
  END              
  Select @Row=Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and              
   (Ins.TdFrBLoc <> ISNULL((SELECT Del.TdFrBLoc FROM Deleted Del WHERE Ins.TdPrtKey=Del.TdPrtKey               
    and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   AND NOT EXISTS (Select 'x' From Loc where LocCoCd= TdCoCd and              
    LocTyp= 'P' and LocCd = Ins.TdFrBLoc AND LocValidYN = 'Y')              
   OR Ins.TdFrBLoc = '')              
   AND @wTcTyp = 'REJ' And TdBDc = 'D'              
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Bag Fr Loc'              
     GoTo ErrHndlr              
  END              
  --5717
  Select @Row=Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and              
   Not Exists (Select 'x' From Loc where LocCoCd= TdCoCd and              
   LocTyp= 'P' and LocCd = Ins.TdFrBLoc)              
    If @Row>0              
	BEGIN              
		Select @ErrMsg='Invalid Bag Fr Loc'              
		GoTo ErrHndlr              
	END    
    end          
              
  if @mode = 'A'              
  Begin              
   select @Row=Count('x') from inserted where @wTcTyp= 'MLT' and
    exists (Select Cast(BQty As Decimal(8,1)) from Bag where TdBIdNo=BIdNo and (cast(Bqty as decimal(8,1))> 0 or BPtQty<>0))
   If @Row > 0              
   begin              
    Select @ErrMsg='Cannot Melt Bag As Bag Qty/Part Qty <> 0'              
    GoTo ErrHndlr              
   end              
              
	
   select @Row=Count('x') from inserted ins join Bag on BIdNo=TdBIdNo               
    where @wTcTyp not in('MLT', 'REJ') and (Round(BRjQty, 1)> 0 or BRjPtQty<> 0)              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Make Any Transaction For A Rejected Bag'              
    GoTo ErrHndlr              
   end              
   --5717
   select @Row=Count('x') from inserted ins where TdPrtKey=@wCurrPrtn and              
    exists (Select 'x' from Txnd Td where Td.TdPrtKey=ins.TdPrtKey and               
     Td.TdCoCd= ins.TdCoCd and Td.TdYy= ins.TdPrevYy and Td.TdKey= ins.TdPrevKey and              
     (Select PValue from Param              
      where PTyp= 'TC' and PMCd= Td.TdTc)= 'MLT')              
   If @Row > 0              
   begin              
    Select @ErrMsg= 'Cannot Make Any Transaction For A Melted Bag'              
    GoTo ErrHndlr              
   end              
 End              
   --5717
   if @wTcTyp in ('BV', 'REJ')              
  begin             /* Current Partition */              
   Select @Row= Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdLotNo<> ''              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Lot No. Should Be Blank'              
      GoTo ErrHndlr              
   end              
  end
  --5717
  select @Row=Count('x') from inserted ins where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and TdLotNo<> ''              
  If @Row > 0              
  begin              
   Select @ErrMsg= 'Lot No Should Be Blank For TdSrNo= 0'              
   GoTo ErrHndlr              
  end 
             
 END      
--5717
 if @mode in ('A','C') and @wEntTyp= 'TxnRm'              
 Begin              
   If @wTcTyp= 'PR'               
   Begin                
  /***** Either All Best Use Order Fields are filled up or all are blank *****/        
  Select @Row= Count('x') From Inserted Ins       
  Where Ins.TdPrtKey=@wCurrPrtn and (TdBstOdTc<> '' or TdBstOdYy<> '' or TdBstOdChr<> '' or TdBstOdNo<> 0 )      
   and (TdBstOdTc= '' or TdBstOdYy= '' or TdBstOdChr= '' or TdBstOdNo= 0 )       
  If @Row> 0      
  begin      
     Select @ErrMsg= 'Either All Order Fields Should Be Filled Or All Should Be Blank'      
     GoTo ErrHndlr       
  end     
  --5717
  Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPRtn and TdBstOdNo<> 0       
   and not exists(Select 'x' From OrdMst where OmPrtKey=TdPrtkey and OmCoCd= TdCoCd and OmTc= TdBstOdTc and OmYy= TdBstOdYy       
    and OmChr= TdBstOdChr and OmNo= TdBstOdNo)       
  If @Row> 0       
  begin      
     Select @ErrMsg= 'Order Number Not Defined'      
     GoTo ErrHndlr      
  end 
  --5717
  Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPRtn and TdBstOdSr <> 0       
   and not exists(Select 'x' From OrdDsg where OdPrtKey=TdPrtkey and OdCoCd= TdCoCd and OdTc= TdBstOdTc and OdYy= TdBstOdYy       
    and OdChr= TdBstOdChr and OdNo= TdBstOdNo and OdSr =TdBstOdSr )       
  If @Row> 0       
  begin      
     Select @ErrMsg= 'Order Serial Number Not Defined'      
     GoTo ErrHndlr      
  end  
  --5717
  Select @Row= Count('x') From Inserted Ins Where TdPrtKey=@wCurrPrtn and TdBstDmCd<> ''       
   and not exists(Select 'x' From DsgMst where DmPrtKey=TdPrtKey and DmTcTyp= 'DM' and DmCd= TdBstDmCd and DmSz= '')       
  If @Row> 0       
  begin      
     Select @ErrMsg= 'Design Code Not Defined'      
     GoTo ErrHndlr      
  end    
  End      
 End 


if @wTcTyp ='BM' and @wEntTyp ='TxnBag'
begin
	select @Row= Count('x') from inserted Ins where              
    Not Exists (Select 'x' from Txnd where Txnd.TdTIdNo=Ins.TdTIdNo              
      and Txnd.TdCoCd= Ins.TdCoCd             
         and Txnd.TdSrNo= 0 and Txnd.TdBDc= 'C' and Txnd.TdRefKey= Ins.TdRefKey and Txnd.TdRefYy =0)               
    and TdBDc= 'D' and TdSrNo= 0              
     If @Row> 0              
        Begin            
    Select @ErrMsg= 'No Parent Bag For The Child Bag'              
      GoTo ErrHndlr              
     end        
	    select @Row= Count('x') from inserted Ins               
    Join Txnd Td On Td.TdPrtKey=Ins.TdPrtKey and Td.TdTIdNo=Ins.TdTIdNo              
     and Td.TdCoCd= Ins.TdCoCd and Td.TdTc= Ins.TdTc and Td.TdYy= Ins.TdYy and Td.TdChr= Ins.TdChr and               
     Td.TdNo= Ins.TdNo and Td.TdSrNo= 0 and Td.TdRefKey= Ins.TdRefKey               
    where Ins.TdSrNo= 0 and               
     not exists (select 'x' from Bag B1               
      join Bag B2 on B1.BPrtKey=B2.BPrtKey and B1.BOdIdNo=B2.BOdIdNo              
       and B1.BCoCd= B2.BCoCd and B1.BOdTc= B2.BOdTc               
       and B1.BOdYy= B2.BOdYy and B1.BOdChr= B2.BOdChr               
       and B1.BOdNo= B2.BOdNo and B1.BOdSr= B2.BOdSr               
      where B1.BPrtKey=Td.TdPrtKey and B1.BIdNo=Td.TdBIdNo and B1.BCoCd= Td.TdCoCd               
       and B1.BYy= Td.TdBYy and B1.BChr= Td.TdBChr and B1.BNo= Td.TdBNo               
       and B2.BCoCd= Ins.TdCoCd and B2.BYy= Ins.TdBYy               
       and B2.BChr= Ins.TdBChr and B2.BNo= Ins.TdBNo)               
        If @Row> 0               
        Begin              
    Select @ErrMsg= 'Order Design Serial Not Same As That Of Previous Bag'               
      GoTo ErrHndlr               
        end              

	    --577 chk -it is commented in 577 trigger, so now in 5715 also commented
		--select @Row= Count('x') from inserted Ins              
  --  Join Txnd Td On Td.TdCoCd= Ins.TdCoCd and              
  --   Td.TdTc= Ins.TdTc and Td.TdYy= Ins.TdYy and Td.TdChr= Ins.TdChr and              
  --   Td.TdNo= Ins.TdNo and Td.TdSrNo= 0              
  --  where Ins.TdSrNo= 0 and              
  --   (select BOdDmCd from Bag where BCoCd= Td.TdCoCd and              
  --    BYy= Td.TdBYy and BChr= Td.TdBChr and BNo= Td.TdBNo)<>              
  --   (select BOdDmCd from Bag where BCoCd= Ins.TdCoCd and              
  --    Byy= Ins.TdBYy and BChr= Ins.TdBChr and BNo= Ins.TdBNo)              
  --      If @Row> 0              
  --      Begin              
  --  Select @ErrMsg= 'Design Code Not Same As That Of Previous Bag'              
  --    GoTo ErrHndlr              
  --      end              

		        select @Row= Count('x') from inserted Ins where              
    Exists (Select 'x' from Txnd Td              
      join Txnd TdRm on TdRm.TdCoCd= Td.TdCoCd              
       and TdRm.TdBYy= Td.TdBYy and TdRm.TdBChr= Td.TdBChr              
       and TdRm.TdBNo= Td.TdBNo and TdRm.TdSrNo<> 0              
      where Td.TdSrNo= 0 and Td.TdBDc= 'C' and Td.TdCoCd= Ins.TdCoCd and              
       Td.TdTc= Ins.TdTc and Td.TdYy= Ins.TdYy and              
       Td.TdChr= Ins.TdChr and Td.TdNo= Ins.TdNo and                     
       Td.TdRefKey= Ins.TdRefKey and               
       TdRm.TdRmCd Not In(select OrRmCd from Bag              
        join OrdRm on BCoCd= OrCoCd and              
          BOdTc= OrTc and BOdYy= OrYy and BOdChr= OrChr              
         and BOdNo= OrNo and BOdSr= OrSr              
         where BCoCd= Ins.TdCoCd and BYy= Ins.TdBYy              
          and BChr= Ins.TdBChr and BNo= Ins.TdBNo)              
         )              
          and Ins.TdSrNo= 0 and Ins.TdBDc= 'D'              
    If @Row> 0              
        Begin              
    Select @ErrMsg= 'RM Code Of Parent Not Same As Child'              
      GoTo ErrHndlr              
        end              

end
 if @wenttyp='TxnBag' 
 begin
  declare @wBLoc varchar(8)=null
  declare @wBRecvDt date=null
  declare @wBMaxTxnDt date=null
  declare @wBQty float 
  declare @wBPtQty int 
  declare @wBGrWt float 
  declare @wBRepQty float 
  declare @wBRjQty float 
  declare @wBRjPtQty int 
  declare @wBRepGrWt float
  declare @wBWrk varchar(8)

 
  If @wTcTyp<> 'BM' and (@mode= 'D' or (@mode= 'C' and (Update (TdBQty) or update (TdPtQty) or update (TdRjQty) or update (TdRjPtQty) ) ) )              
  Begin              
		select @wBQty = (case when TTrfQty=0 then (del.TdBQty* (case del.TdBDc when 'C' then 1 when 'D'              
				    then -1 else 0 end)) 
							 else
							 0 end),
		@wBRepQty= (case when TTrfQty=0 then (del.TdBQty * (case del.TdBDc when 'C' then 1 when 'D'              
					then -1 else 0 end))
							 else
							0 end),
		@wBPtQty = (case when TTrfQty=0 then (del.TdPtQty* (case del.TdBDc when 'C' then 1 when 'D'
					then -1 else 0 end)) else 0 end),              
		@wBRjQty= (case when TTrfQty=0 then (del.TdRjQty* (case del.TdRjDc when 'C' then 1 when 'D'              
					then -1 else 0 end)) else 0 end),              
		@wBRjPtQty= (case when TTrfQty=0 then (del.TdRjPtQty* (case del.TdRjDc when 'C' then 1 when 'D'              
					then -1 else 0 end)) else 0 end)             
	from deleted del join Txn on
	Txn.TIdNo=del.TdTIdNo
		where del.TdSrNo=0                
  End              

              
  If @wTcTyp= 'BM' and (@mode= 'D' or (@mode= 'C' and (Update (TdBQty) or update (TdPtQty) or update (TdRjQty) or update (TdRjPtQty) ) ) )              
  Begin              
   select 
   @wBQty= del.TdBQty * (case del.TdBDc when 'C' then 1 when 'D'              
    then 0 else 0 end),              
   @wBPtQty=  del.TdPtQty* (case del.TdBDc when 'C' then 1 when 'D'              
    then 0 else 0 end),              
   @wBRjQty= del.TdRjQty* (case del.TdRjDc when 'C' then 1 when 'D'              
    then 0 else 0 end),              
   @wBRjPtQty= del.TdRjPtQty* (case del.TdRjDc when 'C' then 1 when 'D'              
    then 0 else 0 end)              
   from deleted del join Param on PTyp= 'BCHR' and              
     PMCd= del.TdBChr and PSCd= '' and PValue in ('C', 'F') 
    and del.TdSrNo=0                
  End          
              
  select 
  @wBGrWt = (((case (case when del.TdFrRmLoc='W' then del.TdFrRmDc              
            when del.TdToRmLoc='W' then del.TdToRmDc end)              
                     when 'C' then 1 when 'D' then -1 else 0 end)* del.TdRmWt/              
   (case when tdRmCtg in ('D','C') then 5 else 1 end))),   
 
  @wBRepGrWt = (((case (case when del.TdFrRmLoc='XR' then del.TdFrRmDc              
            when del.TdToRmLoc='XR' then del.TdToRmDc end)              
                     when 'C' then 1 when 'D' then -1 else 0 end)* del.TdRmWt/              
   (case when tdRmCtg in ('D','C') then 5 else 1 end)))   
  from deleted del where del.TdSrNo<>0 and
   del.TdByy<> '' and del.TdBChr<> ''              
   and del.TdBNo<> 0 and del.TdRmCd<> ''              


 if @mode in ('D') and @wEntTyp= 'TxnBag' and (select TdSrNo from deleted)=0
 begin
  declare @wFgToBLoc varchar(8)
  declare @wTdYyKeyAckBV varchar(10)
  
  Select @wFgToBLoc = Isnull(F.FgToBLoc,'') 
	from Fgd Fd              
	join Deleted Del on FdBIdNo=Del.TdBIdNo
    join Fg F on F.FgIdNo=Fd.FdFgIdNo 
    where Fd.FdCoCd=Del.tdcocd and Del.TdSrNo=0 
     and Fd.FdBYyKey=Del.TdPrevYy + substring(str(Del.TdPrevKey),3,8)
     and FdTc='FWB' 

   Select 
	@wTdYyKeyAckBV=
	max(td.TdYyKey) 
     from Txnd td join Deleted Del
     on del.TdBIdNo=td.TdBIdNo               
     and td.TdTc='BV' and Td.TdAck='Y'
  

   Select @wBRecvDt = isnull(TDt,BOpnDt), 
		  @wBLoc= isnull(@wFgToBLoc,Txnd.TdToBLoc)
	  From Txnd 
	  join Txn on TdTIdNo=TIdNo
	  join Bag on BIdNo=Txnd.TdBIdNo 
	  join Deleted TdDel on TdDel.TdBIdNo=Txnd.TdBIdNo
	  where Txnd.TdYyKey=@wTdYyKeyAckBV 

   if @wBLoc is null
   begin
   Select @wBLoc=case when BRepYN='N' then BOpnLoc 
	else
	(select HDefRepBagLoc from Head where HCoCd=Deleted.TdCocd) end
   from Deleted, Bag where BIdNo=Deleted.TdBIdNo
   end
  End 

  If (@mode='D')
  Begin    
       select @wBMaxTxnDt = IsNull((Select Max(qDt)               
           From (Select IsNull(Max(Tddt),'01/01/80') qDt               
           From Txnd               
           where TdBIdNo=Del.TdBIdNo and TdIdNo<>del.TdIdNo                
        Union All              
                Select IsNull(Max(Fddt),'01/01/80') qDt
           From Fgd
           where FdBIdNo=Del.TdBIdNo 
		   Union All              
            Select BOpnDt qDt From Bag               
           Where BIdNo=Del.TdBIdNo ) B),null)              
   From	 Deleted Del              


  End

   Update Bag set BLoc=case when BRepYN='N' and @wBLoc is not null then @wBLoc else BLoc  end,  
				  BRecvDt = isnull(@wBRecvDt ,BRecvDt), 
				  BRepLoc = case when BRepYN='Y' and @wBLoc is not null then @wBLoc else BRepLoc end,
				  BWrk = (case when deleted.TdSrNo=0 then isnull(Txnd.TdWrk,'ZSELF') else BWrk end),
				  BLstYy =(case when @mode='D' and deleted.TdSrNo=0 then isnull(Txnd.TdYy,'') else BLstYy end),
				  BLstKey=(case when @mode='D' and deleted.TdSrNo=0 then isnull(Txnd.TdKey,0) else BLstKey end),
 				  BLstYyKey=(case when @mode='D' and deleted.TdSrNo=0 then isnull(Txnd.TdYyKey,'') else BLstYyKey end),				  
				  BMaxTxnDt = (case when deleted.TdSrNo=0 then isnull(@wBMaxTxnDt ,BOpnDt) else  BMaxTxnDt end),
				  BQty=BQty+ (case when deleted.TdSrNo=0 then case when BRepYN='N' then isnull(@wBQty,0) else 0 end else 0 end),
				  BRepQty=BRepQty + (case when deleted.TdSrNo=0 then case when BRepYN='Y' then isnull(@wBRepQty,0) else 0 end else 0 end),
				  BPtQty = BPtQty+(case when deleted.TdSrNo=0 then isnull(@wbptqty,0) else 0 end),
				  bRjQty=bRjQty+isnull(@wbRjQty,0),
				  bRjPtQty=bRjPtQty+isnull(@wBRjPtQty,0)
		from Bag 
		join Deleted on deleted.TdBIdNo=Bag.BIdNo
		left outer join Txnd on 
		deleted.TdBIdNo=Bag.BIdNo and Txnd.TdCoCd=deleted.TdCoCd and Txnd.TdYy = deleted.TdPrevYy and Txnd.TdKey=deleted.TdPrevKey 
		where Deleted.TdSrNo=0 


  Update Bag set              
  BGrWt= BGrWt+ (case when BRepYN ='N' then IsNull((Select sum((case (case when del.TdFrRmLoc='W' then del.TdFrRmDc              
            when del.TdToRmLoc='W' then del.TdToRmDc end)              
                     when 'C' then 1 when 'D' then -1 else 0 end)* del.TdRmWt/              
   (case when RmCtg in ('D','C') then 5 else 1 end))              
   from deleted del Join RmMst on RmCd= del.TdRmCd 
   where del.TdBIdNo=BIdNo and del.TdRmCd<> ''), 0) else 0 end),
  BRepGrWt= BRepGrWt+ (case when BRepYN ='Y' then IsNull((Select sum((case (case when del.TdFrRmLoc='XR' then del.TdFrRmDc              
            when del.TdToRmLoc='XR' then del.TdToRmDc end)              
                     when 'C' then 1 when 'D' then -1 else 0 end)* del.TdRmWt/              
   (case when RmCtg in ('D','C') then 5 else 1 end))              
   from deleted del Join RmMst on RmCd= del.TdRmCd where del.TdBIdNo=BIdNo and del.TdRmCd<> ''), 0) else 0 end)
  from deleted where BIdNo=deleted.TdBIdNo and deleted.TdRmCd<> ''              


  If @wEntTyp= 'TxnBag' and @wTcTyp= 'BM'              
  Begin              
   If @Mode in ('D')
   Begin              
    Update Bag               
     Set BCls = 'N',              
         BClsDt = '01/01/80'              
    From Deleted Del              
    Where BIdNo=Del.TdBIdNo and BCoCd= Del.TdCoCd And BYy= Del.TdBYy And BChr= Del.TdBChr And BNo= Del.TdBNo               
--     And Convert(Decimal(16, 3), BGrWt) <> Convert(Decimal(16, 3), 0)               
     And (Select PValue From Param Where PTyp= 'BCHR' And PMCd= Del.TdBChr And PSCd= '') in ('C', 'F')               
   End              
  End        
  
  If @wEntTyp= 'TxnBag' and @wTcTyp= 'MLT'              
  Begin              
   If @Mode in ( 'D' ,'C')             
   Begin              
    Update Bag               
     Set BCls = 'N',              
         BClsDt = '01/01/80'              
    From Deleted Del              
    Where BIdNo=Del.TdBIdNo and BCoCd= Del.TdCoCd And BYy= Del.TdBYy And BChr= Del.TdBChr And BNo= Del.TdBNo               
--     And Convert(Decimal(16, 3), BGrWt) <> Convert(Decimal(16, 3), 0)               
   End              
  End              

 End              

        
 if @mode in ('A','C') and @wEntTyp= 'TxnBag'              
 begin              
  If (@mode in( 'A','C'))              
  Begin              
   Select @wBMaxTxnDt = TdDt              
   From Inserted                  
    Join Bag on BIdNo=TdBIdNo 
   Where TdDt>=BMaxTxnDt                 
  End              
              
  select @wBWrk = case when ins.TdWrk<>'' and ins.TdSrNo=0 then ins.TdWrk else null end
  from inserted ins 

  If @wTcTyp<> 'BM' and (@mode= 'A' or (@mode= 'C' and (Update (TdBQty) or update (TdPtQty) or update (TdRjQty) or update (TdRjPtQty) or update(TdYy) or update (TdKey) or update (TdYyKey) ) ) )              
  Begin              
   select 
   @wBQty= case when TTrfQty=0 then (ins.TdBQty* (case ins.TdBDc when 'C' then -1 when 'D'              
    then 1 else 0 end)) else 0 end,              
   --BRepQty= BRepQty+ case when BRepYN ='Y' then (ins.TdBQty* (case ins.TdBDc when 'C' then -1 when 'D'              
   -- then 1 else 0 end)),              
   @wBPtQty =  case when TTrfQty=0 then (ins.TdPtQty* (case ins.TdBDc when 'C' then -1 when 'D'              
    then 1 else 0 end)) else 0 end,              
   @wBRjQty= case when TTrfQty=0 then (ins.TdRjQty* (case ins.TdRjDc when 'C' then -1 when 'D'              
    then 1 else 0 end)) else 0 end ,              
   @wBRjPtQty= case when TTrfQty=0 then (ins.TdRjPtQty* (case ins.TdRjDc when 'C' then -1 when 'D'              
    then 1 else 0 end)) else 0 end
 from inserted ins join Txn on Txn.TIdNo=ins.TdTIdNo
	join Param on PTyp= 'BCHR' and              
      PMCd= ins.TdBChr and PSCd= '' 
   where ins.TdSrNo=0         

  End              
              
  If @wTcTyp= 'BM'               
  Begin              
   If (@mode= 'A' or (@mode= 'C' and (Update (TdBQty) or update (TdPtQty) or update (TdRjQty) or update (TdRjPtQty)  ) ) )              
   Begin              
    select @wBQty=
     ((ins.TdBQty* (case ins.TdBDc when 'C' then -1 when 'D'              
     then 0 else 0 end))) ,              
    @wBPtQty=((ins.TdPtQty* (case ins.TdBDc when 'C' then -1 when 'D'              
     then 0 else 0 end)))              
    from inserted ins where ins.TdSrNo= 0 and (Select PValue from Param where PTyp= 'BCHR' and              
      PMCd= ins.TdBChr and PSCd= '' ) in ('C', 'F')               
   End               
                 
  End              
              
  Select @wBLoc= ins.TdToBLoc ,              
   @wBRecvDt=ins.TdDt 
   from inserted ins where 
    ins.TdToBLoc<> '' and ins.TdSrNo= 0 and ins.TdAck= 'Y'              
              
 
   update Bag set 
	   BMaxTxnDt = @wBMaxTxnDt,
	   BLoc=case when BRepYN='N' and @wBLoc is not null then @wBLoc else BLoc end,
	   BRecvDt=isnull(@wBRecvDt, BRecvDt),
	   --5716 repair bv issue- BRepLoc=case when BRepYn='N' and @wBLoc is not null then @wBLoc else BRepLoc end,
	   BRepLoc=case when BRepYn='Y' and @wBLoc is not null then @wBLoc else BRepLoc end,
	   BWrk=isnull(@wBWrk,Bwrk),
       BLstYy= TdYy ,              
       BLstKey= convert(int, substring((ins.TdYyKey),3, 8))              ,
       BLstYyKey= (ins.TdYyKey), 
	   BQty=Bqty+case when BRepYN='N' then isnull(@wBQty,0) else 0 end,
	   BRepQty=BRepQty+case when BRepYN='N' then isnull(@wBRepQty,0) else 0 end,
	   BGrWt =BGrWt+case when BRepYN='N' then isnull(@wBGrWt,0) else 0 end,
	   BPtQty=BPtQty+case when BRepYN='N' then isnull(@wBPtQty,0) else 0 end,
	   BRepGrWt =BRepGrWt+case when BRepYN='Y' then isnull(@wBRepGrWt,0) else 0 end,
	   BRjQty=bRjQty+isnull(@wBRjQty,0),
       BRjPtQty=bRjPtQty+isnull(@wBRjPtQty,0)
   from inserted ins where BIdNo=ins.TdBIdNo          and ins.tdsrno=0


  update Bag set              
  BMaxTxnDt=isnull(@wBMaxTxnDt,BOpnDt),
  BGrWt= BGrWt+ (case when BRepYN='N' then IsNull((Select ((case (case when ins.TdFrRmLoc='W' then ins.TdFrRmDc              
          when ins.TdToRmLoc='W' then ins.TdToRmDc end)              
          when 'C' then -1 when 'D' then +1 else 0 end)              
   * ins.TdRmWt/ (case when RmCtg in ('D','C') then 5 else 1 end))              
   from inserted ins Join RmMst on RmCd= ins.TdRmCd		
   where ins.TdBIdNo=BIdNo and ins.TdCoCd= BCoCd and ins.TdBYy= BYy and ins.TdBChr= BChr              
   and ins.TdBNo= BNo and ins.TdRmCd<> ''), 0) else 0 end),             
  BRepGrWt= BRepGrWt+ (case when BRepYN='Y' then IsNull((Select ((case (case when ins.TdFrRmLoc='XR' then ins.TdFrRmDc              
          when ins.TdToRmLoc='XR' then ins.TdToRmDc end)              
          when 'C' then -1 when 'D' then +1 else 0 end)              
   * ins.TdRmWt/ (case when RmCtg in ('D','C') then 5 else 1 end))              
   from inserted ins Join RmMst on RmCd= ins.TdRmCd 
   where ins.TdBIdNo=BIdNo and ins.TdRmCd<> ''), 0) else 0 end)
  from inserted where BIdNo=inserted.TdBIdNo and inserted.TdRmCd<> ''              

 end              
 
 IF @wEntTyp='TxnBag'              
 Begin              
  Select @Row= Count('x') From Inserted Ins Where 
   Not Exists (Select 'x' From Bag Where BIdNo=Ins.TdBIdNo and BCoCd= Ins.TdCoCd and BYy= Ins.TdByy              
     and BChr= Ins.TdBChr and BNo= Ins.TdBNo and (BCls= 'N' or BRepYN ='Y'))              
     If @Row>0              
  begin              
   Select @ErrMsg='Invalid Bag No/Bag Is Closed'              
     GoTo ErrHndlr              
  end              
  --5717
    select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and @wTcTyp              
   in('DT', 'BV', 'MLT') and TdPrtKey=@wCurrPrtn and TdSrNo= 0 and not exists(Select 'x' from Txn where              
    TPrtKey=TdPrtKey and TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr and TNo= TdNo and              
   TFrBLoc= (case when TdBDc= 'C' then TdFrBLoc when TdBDc= 'D' then TdToBLoc else TdFrBLoc end))              
   and IsNull((Select hOthLocBag from Head where HCoCd= TdCoCd and HCd= @wHCd), '') = 'N'              
    if @row> 0              
  begin              
      select @errmsg= 'From Bag Loc Should Be Equal To The Default Bag Location'              
   goto ErrHndlr              
    end 

      select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and @wTcTyp               
   in('BS', 'BM') and TdPrtKey=@wCurrPrtn and TdSrNo= 0 and not exists(Select 'x' from Txn where               
    TPrtKey=TdPrtKey and TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc and TYy= TdYy and TChr= TdChr               
    and TNo= TdNo and TFrBLoc= TdFrBLoc)               
   and IsNull((Select hOthLocBag from Head where HCoCd= TdCoCd and HCd= @wHCd), '') = 'N'               
    if @row> 0               
  begin               
      select @errmsg= 'From Bag Loc Should Be Equal To The Default Bag Location'               
   goto ErrHndlr               
    end              
  /* **** Zubin 211 **** */              
  
     Select @Row= Count('x') From Inserted Ins Where Ins.TdPrtKey=@wCurrPrtn and              
   Not Exists (Select 'x' From Param where PTyp='DC'              
        and PMCd = Ins.TdBDc and PScd='')              
   and TdBDc<>''              
    If @Row>0              
  Begin              
   Select @ErrMsg='Invalid Bag Dc'              
   GoTo ErrHndlr              
  End              
  select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and @wTcTyp               
   in('BM') and TdPrtKey=@wCurrPrtn and ((Select count(distinct TdBChr) from Txnd where TdCoCd = inserted.TdCoCd and TdTc = inserted.TdTc and TdYy = Inserted.TdYy and TdChr = inserted.TdChr 
									and TdNo = inserted.TdNo and TdSrNo = 0 and TdBDc='D' and tdrefkey = inserted.tdrefkey)>1 
									or (Select count(distinct TdBChr) from Txnd where TdCoCd = inserted.TdCoCd and TdTc = inserted.TdTc and TdYy = Inserted.TdYy and TdChr = inserted.TdChr 
									and TdNo = inserted.TdNo and TdSrNo = 0 and TdBDc='C' and tdrefkey = inserted.tdrefkey)>1)  
                
    if @row> 0               
  begin               
      select @errmsg= 'Issue CHR must be one Receive CHR must be one'               
   goto ErrHndlr               
    end              
	
   select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and @wTcTyp               
   in('BM') and TdPrtKey=@wCurrPrtn and (Select max(distinct TdBChr) from Txnd where TdCoCd = inserted.TdCoCd and TdTc = inserted.TdTc and TdYy = Inserted.TdYy and TdChr = inserted.TdChr 
									and TdNo = inserted.TdNo and TdSrNo = 0 and TdBDc='D' and tdrefkey=inserted.tdrefkey)=
									(Select max(distinct TdBChr) from Txnd where TdCoCd = inserted.TdCoCd and TdTc = inserted.TdTc and TdYy = Inserted.TdYy and TdChr = inserted.TdChr 
									and TdNo = inserted.TdNo and TdSrNo = 0 and TdBDc='C' and tdrefkey=inserted.tdrefkey) 
                
    if @row> 0               
  begin               
      select @errmsg= 'Issue CHR must be different from Recv Chr'               
   goto ErrHndlr               
    end              

  /*********** TdBGrWt >=0 **********/              
  /* Current Partition */              
  Select @Row=Count('x') From Inserted Ins Where TdBGrWt< 0              
  If @Row>0              
    Begin              
   Select @ErrMsg='Gross Wt cannot be -ve'              
     GoTo ErrHndlr              
    End              
              
--/*  ######################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ######################  */              
  Select @Row=Count('x') From Inserted Ins              
   Join Bag  On BIdNo=TdBIdNo 
  Where (TdSrNo<> 0 and Not Exists (Select 'x' from OrdRm where              
	OrOdIdNo=BOdIdNo  and OrRmCd= TdRmCd))              
   OR              
   (TdSrNo=0 and TdRmCd<> '')              
    If @Row> 0              
  BEGIN              
   Select @ErrMsg='Rm Code does not exist in Order/Rm Code Should Be Blank'              
   GoTo ErrHndlr              
  END              
 
  --SELECT @Row= COUNT('x') FROM Inserted Ins WHERE (TdSrNo<>0 AND              
  -- NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
  --    LocTyp<> 'P' and LocTyp <>'R' AND LocCd = Ins.TdToRmLoc AND LocValidYN='Y' And              
  --     (@wTcTyp NOT IN('Rec','Cnv','Bs', 'BM') OR LocTyp= 'Xc')))              
  -- OR              
  --  (TdSrNo=0 AND TdToRmLoc<>'')              
  --  IF @Row> 0              
  --BEGIN              
  -- SELECT @ErrMsg= 'Invalid To Rm Loc/ To Rm Loc Should Be Blank'              
  -- GOTO ErrHndlr              
  --END              
               
  select @row= count('x') from inserted Ins where TdSrNo<>0 and              
    Ins.TdFrRmLoc = Ins.TdToRmLoc and (Ins.TdFrRmLoc <> '' or              
    Ins.TdToRmLoc <> '')              
    if @row>0            
    begin              
      select @errmsg= 'To Rm Loc and From Rm Loc cannot be same'              
      goto ErrHndlr              
    end              

   if @wTcTyp='BS'
   begin
	  Select @Row= Count('x') from inserted Ins where TdSrNo= 0 and TdBDc= 'D' and               
       Not Exists (Select 'x' from Txnd                
      where Txnd.TdTIdNo=Ins.TdTIdNo              
		and         Txnd.TdSrno= 0 and Txnd.TdBDc= 'C')               
     If @Row > 0               
     Begin              
    Select @ErrMsg= 'No Parent Bag For The Child Bag '              
      GoTo ErrHndlr              
     end          
	    select @Row= Count('x') from inserted Ins       
  Join Txn On TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where TMultiIr='R' and       
    (Select count('x') from Txnd  where 
     Txnd.TdTIdNo=Ins.TdTIdNo and              
     Txnd.TdSrNo= 0 and Txnd.TdBDc='C')> 1              
    and TdBDc='C' and TdSrNo= 0              
   If @Row > 0              
      begin              
	 Select @ErrMsg='Only One Parent Bag allowed in Bag Split'              
    GoTo ErrHndlr              
   end              
    
    select @Row= Count('x') from inserted Ins       
  Join Txn On TIdNo=Ins.TdTIdNo   where 
    Not Exists (Select 'x' from Txnd where Txnd.TdCoCd= Ins.TdCoCd and            
         ((TMultiIr='R' and Txnd.TdYy= Ins.TdRefYy and Txnd.TdKey= Ins.TdRefKey)      
     or       
         (TMultiIr='I' and Txnd.TdCoCd=Ins.TdCoCd and Txnd.TdRefYy= Ins.TdYy and Txnd.TdRefKey= Ins.TdKey))      
         and (select BOdDmCd from Bag where BIdNo=Txnd.TdBIdNo              
       )              
           = (select BOdDmCd from Bag where BIdNo=Ins.TdBIdNo              
       )              
            )              
          and TdBDc= 'D' and TdSrNo= 0              
    and Str(TtrfQty, 16, 1) = Str(0, 16, 1)              
        If @Row > 0              
        Begin              
    Select @ErrMsg='Design Code Of Parent Not Same As Child '      
      GoTo ErrHndlr              
        end             
       

 If Exists(Select 'x' from Txn      
   Join Inserted on TdCoCd=TCoCd and TTc=TdTc and TYy=TdYy and TChr=TdChr and TNo=TdNo and       
    TIdNo=TdTIdNo 
   Where TdBDc= 'D' and TdSrNo= 0 and TMultiIr='I') and @mode='A'      
 Begin           
    select @Row= Count('x') from inserted Ins       
  where 
    Exists (Select 'x' from Txnd Td               
 join Txnd Tdr on Tdr.TdBIdNo=Td.TdBIdNo              
         and Tdr.TdSrNo<> 0               
     where Td.TdSrNo= 0 and Td.TdCoCd= Ins.TdCoCd and           
   Td.TdRefYy= Ins.TdYy and Td.TdRefKey= Ins.TdKey      
         and Tdr.TdRmCd Not in (select OrRmCd from Bag               
        join OrdRm on BOdIdNo=OrOdIdNo 
         where BIdNo=Ins.TdBIdNo)               
         )              
          and Ins.TdBDc= 'D' and Ins.TdSrNo= 0              
    If @Row > 0              
        Begin              
    Select @ErrMsg= 'RM Code Of Parent Not Same As Child'              
      GoTo ErrHndlr              
        end              
 END      
      
  /***** Bef 3.8.0 *****/      
 If Exists(Select 'x' from Txn      
   Join Inserted on TdCoCd=TCoCd and TTc=TdTc and TYy=TdYy and TChr=TdChr and TNo=TdNo and       
    TIdNo=TdTIdNo 
   Where TdBDc= 'D' and TdSrNo= 0 and TMultiIr='R')  and @mode='A'      
 Begin      
    select @Row= Count('x') from inserted Ins where 
    Exists (Select 'x' from Txnd Td               
       join Txnd Tdr on Tdr.TdBIdNo=Td.TdBIdNo              
				and Tdr.TdSrNo<> 0               
     where Td.TdSrNo= 0 and Td.TdCoCd= Ins.TdCoCd and               
         Td.TdYy= Ins.TdRefYy and Td.TdKey= Ins.TdRefKey               
         and Tdr.TdRmCd not in(select OrRmCd from Bag               
        join OrdRm on BOdIdNo=OrOdIdNo 
         where BIdNo=Ins.TdBIdNo)               
         )              
          and Ins.TdBDc= 'D' and Ins.TdSrNo= 0              
    If @Row > 0              
        Begin              
    Select @ErrMsg= 'RM Code Of Parent Not Same As Child'              
      GoTo ErrHndlr              
        end       
	Select @Row= Count('x') from inserted Ins where 
    exists (Select 'x' from Txn where 
      TIdNo=Ins.TdTIdNo 
and         Str(TtrfQty, 16, 1) <> Str(0, 16, 1)) and              
    Exists (Select 'x' from Txnd TdD where TdD.TdSrNo<> 0 
      and TdD.TdBIdNo=Ins.TdBIdNo 
    and TdD.TdRmCd not in(select OrRmCd from Txnd TdC              
          join Bag on BIdNo=TdC.TdBIdNo              
          join OrdRm on BOdIdNo=OrOdIdNo              
          where TdC.TdSrNo= 0 
         and TdC.TdCoCd= Ins.TdCoCd and TdC.TdYy= Ins.TdRefYy              
         and TdC.TdKey= Ins.TdRefKey)              
         )              
          and Ins.TdBDc= 'D' and Ins.TdSrNo= 0              
        If @Row > 0              
        Begin              
    Select @ErrMsg= 'RM Code Of Child Not Same As Parent'      
      GoTo ErrHndlr              
        end           

 End           
 
 select @Row= Count('x') from inserted Ins       
  Join Txn On TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where TMultiIr='I' and       
    (Select count('x') from Txnd  where 
     Txnd.TdTIdNo=Ins.TdTIdNo and              
     Txnd.TdSrNo= 0 and Txnd.TdBDc='D')> 1              
    and TdBDc='D' and TdSrNo= 0              
   If @Row > 0              
      begin              
    Select @ErrMsg='Only One Child Bag allowed in Bag Split'              
    GoTo ErrHndlr              
   end              

   
   select @Row= Count('x') from inserted Ins       
  Join Txn On TPrtKey=Ins.TdPRtKey and TIdNo=Ins.TdTIdNo and TCoCd=Ins.TdCoCd and TTc=Ins.TdTc and TYy=Ins.TdYy and              
     TChr=Ins.TdChr and TNo=Ins.TdNo      
  where Ins.TdPRtKey=@wCurrPrtn and TMultiIr='I' and       
    exists(Select 'x' from Txnd  where Txnd.TdPrtKey=Ins.TdPRtKey               
     and Txnd.TdTIdNo=Ins.TdTIdNo and Txnd.TdCoCd=Ins.TdCoCd and              
     Txnd.TdTc=Ins.TdTc and Txnd.TdYy=Ins.TdYy and              
     Txnd.TdChr=Ins.TdChr and Txnd.TdNo=Ins.TdNo and              
     Txnd.TdSr<Ins.TdSr and Txnd.TdSrNo= 0 and Txnd.TdBDc='D')      
    and TdBDc='C' and TdSrNo= 0              
   If @Row > 0              
      begin              
    Select @ErrMsg='Only 1 set allowed in Bag Split'              
    GoTo ErrHndlr              
   end
   end 

     select @row= count('x') from inserted Ins where TdPRtKey=@wCurrPrtn and TdSrNo<>0 and              
    Ins.TdFrRmLoc = Ins.TdToRmLoc and (Ins.TdFrRmLoc <> '' or              
    Ins.TdToRmLoc <> '')              
    if @row>0            
    begin              
      select @errmsg= 'To Rm Loc and From Rm Loc cannot be same'              
      goto ErrHndlr              
    end              
              
  /****** if 'TxnBag' and SrNo<> 0 then TdFrRmLoc & TdToRmLoc both should not be of LocTyp= 'W' **********/              
  /* Current Partition */              
  Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and  TdSrNo<>0 and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'W' and LocCd = Ins.TdFrRmLoc)              
     and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'W' and LocCd = Ins.TdToRmLoc)              
    If @Row>0              
  begin              
   Select @ErrMsg= 'Both From & To Rm Locs Should Not Be ''W'' '              
   GoTo ErrHndlr              
  end              
              
  Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and  TdSrNo<>0 and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'XR' and LocCd = Ins.TdFrRmLoc)              
     and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'XR' and LocCd = Ins.TdToRmLoc)              
    If @Row>0              
  begin              
   Select @ErrMsg= 'Both From & To Rm Locs Should Not Be ''XR'' '              
   GoTo ErrHndlr              
  end              


   if @wTcTyp='REJ'
   begin
   select @Row=Count('x') from inserted join Bag on BIdNo=TdBIdNo where 
   (cast(BQty As Decimal(8,2)) <> 0 or              
    (BPtQty <> 0              )
   ) and              
   ( cast(BRjQty As Decimal(8,2)) <> 0 or              
    (BRjPtQty)<> 0              
   )              
  If @Row > 0              
  begin              
   Select @ErrMsg= 'All the Pieces/ Part Pieces In The Bag Have To Be Rejected/ Derejected'              
   GoTo ErrHndlr              
  end   

  Select @Row=Count('x') From Inserted Ins              
   Where (Not Exists (Select 'x' From Param where PTyp='REJCD'  
    and PMCd = Ins.TdRjCd and PScd='')              
    and @wTcTyp='Rej' and TdSrNo=0)              
   OR              
    ((@wTcTyp<>'REJ' or TdSrNo<>0) and TdRjCd<>'')              
    If @Row> 0              
  BEGIN              
   Select @ErrMsg='Invalid Rej Cd/Rej Cd Should Not Be Entered'              
   GoTo ErrHndlr              
  END              
  End          
	-- doubt
  SELECT @Row= COUNT('x') FROM Inserted Ins WHERE (TdSrNo=0 AND              
   ((Ins.TdWrk <> ISNULL((SELECT Del.TdWrk FROM Deleted Del WHERE 
    Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   AND              
   NOT EXISTS (SELECT 'x' FROM vParam WHERE vPCoCd = TdCoCd              
    AND vPTyp='WORK' AND vPMCd = Ins.TdWrk AND vPValidYN = 'Y') OR TdWrk = '')))              
   OR              
   (TdSrNo<>0 AND TdWrk<>'')              
      IF @Row>0              
  BEGIN              
   SELECT @ErrMsg='Invalid Wrk/Worker Should Not Be Entered'              
   GOTO ErrHndlr              
  END              
              

  Select @Row= Count('x') from Deleted Join Bag on BOdIdNo=deleted.TdBIdNo and @wTChr ='Y' where 
    Exists (Select 'x' from 
						OrdDsg  where OrdDsg.OdDmStkCoCd=Deleted.TdCoCd and OrdDsg.OdDmStkYy = Deleted.TdBYy and OrdDsg.OdDmStkChr = Deleted.TdBChr and OrdDsg.OdDmStkNo = Deleted.TdBNo and 
								OrdDsg.OdTc = 'JRO' and (Deleted.ModDt < OrdDsg.ModDt or (Deleted.ModDt = OrdDsg.ModDt and Deleted.ModTime < OrdDsg.ModTime))
			)

    If @Row > 0              
         Begin              
		    Select @ErrMsg= 'Cannot Delete JRO transaction exists for this Bag'              
			GoTo ErrHndlr              
         End              
              
    
   Select @Row= Count('x') From Deleted Del Where Del.TdRmCd<> '' and              
    exists (Select 'a' from Txnd Td where Td.TdBIdNo=Del.TdBIdNo 
	and       Td.TdRmCd= Del.TdRmCd and Td.TdLotNo= Del.TdLotNo              
      and Str(Td.TdRmSz, 16, 4)= Str(Del.TdRmSz, 16, 4)              
      and Str(Td.TdRmSz2, 16, 4)= Str(Del.TdRmSz2, 16, 4)              
      and Str(Td.TdRmSz3, 16, 4)= Str(Del.TdRmSz3, 16, 4)              
      and Str(Td.TdRmStkRt, 16, 3)= Str(Del.TdRmStkRt, 16, 3) having 
				round(sum(Td.TdRmWt* (              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end)),3) < 0 or round(sum(Td.TdRmQty*(              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end)),3) < 0 )
     If @Row>0              
   begin              
    Select @ErrMsg='Rm Wt/Qty Is Less Than Zero In Bag'              
    GoTo ErrHndlr              
   end              
    
   Select @Row= Count('x') From Inserted Ins Where Ins.TdRmCd<> '' and              
    exists (Select 'a' from Txnd Td where Td.TdBIdNo=Ins.TdBIdNo 
	and       Td.TdRmCd= Ins.TdRmCd and Td.TdLotNo= Ins.TdLotNo              
      and Str(Td.TdRmSz, 16, 4)= Str(Ins.TdRmSz, 16, 4)              
      and Str(Td.TdRmSz2, 16, 4)= Str(Ins.TdRmSz2, 16, 4)              
      and Str(Td.TdRmSz3, 16, 4)= Str(Ins.TdRmSz3, 16, 4)              
      and Str(Td.TdRmStkRt, 16, 3)= Str(Ins.TdRmStkRt, 16, 3) having 
				round(sum(Td.TdRmWt* (              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end)),3) < 0 or round(sum(Td.TdRmQty*(              
         Case when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'D' then 1              
      when Td.TdFrRmLoc= 'W' and Td.TdFrRmDc= 'C' then -1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'D' then 1              
      when Td.TdToRmLoc= 'W' and Td.TdToRmDc= 'C' then -1              
             Else 0 end)),3) < 0 )
     If @Row>0              
   begin              
    Select @ErrMsg='Rm Wt/Qty Is Less Than Zero In Bag'              
    GoTo ErrHndlr              
   end              
  
              
  If @wTcTyp= 'BS'              
  Begin              
    /* Current Partition */              
        select @Row= Count('x') from inserted              
    join Txn on TCoCd = TdCoCd and TTc = TdTc and TYy = TdYy              
    and TChr = TdChr and TNo = TdNo              
    where TtrfQty = 1 and TdSrNo=0 and
    (TdBQty = 0 or                
        Round(TdBQty, 1) > Round(IsNull((Select BQty from Bag           
        where BIdNo=TdBIdNo ), 0), 1))              
        If @Row > 0              
        Begin              
    Select @ErrMsg= 'Transfer Quantity cannot be zero or more than Bag Quantity'              
      GoTo ErrHndlr              
        end        
   if @mode in ('D')              
   begin              
              
    /****** Cannot Delete 'Bs', TdDc='C', Tdsrno= 0              
    If exists TdBDc='D', TdSrno=0 in same voucher ********/              
    /* Current Partition */               
    select @Row=Count('x') from Deleted Del join Txn on Txn.TIdNo = del.TdTIdNo where 
     Exists (
			 Select 'x' from Txnd where 
      Txnd.TdTIdNo=Del.TdTIdNo and Txnd.TdBDc= 'D' and TdSrNo= 0              
			)              
           and TdBDc= 'C' and TdSrNo= 0 and TTrfQty=0             
    If @Row > 0              
         begin              
       Select @ErrMsg='Cannot Delete Parent, Child Exists'              
       GoTo ErrHndlr              
         end              
              
   end              
              
  End              
              
              
--/* **** Zubin 211 **** */ /* **** Zubin 211 **** */ /* **** Zubin 211 **** */ /* **** Zubin 211 **** */              
  If @wTcTyp= 'BM'              
  Begin              
   /*** 18/09/2006 ***/              
   /* Current Partition */               
 --  select * from inserted where 
 --  exists (select sum(TdBQty) from Txnd ins where 
	--TdTIdNo=ins.TdTIdNo and TdRefyy=ins.TdRefYY and TdRefKey=ins.TdRefKey
	--and Txnd.TdBDC='C' and Txnd.TdSRNo=0 
	--group by TdBChr
	--having min(qqty) <> max(qqty))


   select @Row= Count('x') from               
     (Select sum(TdBQty) as qQty from Txnd Tc               
     where exists (Select 'x' from inserted ins               
      where ins.TdBDc= 'D' and ins.TdSrNo= 0               
      and ins.TdTIdNo=Tc.TdTIdNo               
      and ins.TdCoCd= Tc.TdCoCd and ins.TdTc= Tc.TdTc               
      and ins.TdYy= Tc.TdYy and ins.TdChr= Tc.TdChr               
	  and ins.TdRefYy=0 
      and ins.TdNo= Tc.TdNo and ins.TdRefKey= Tc.TdRefKey)              
      and Tc.TdSrNo= 0 and Tc.TdBDc= 'C'                
     group By TdBChr) Tbl               
     having convert (decimal (16, 1), min(qQty))<> convert (decimal (16, 1), max(qQty))               
     If @Row> 0               
        Begin              
    Select @ErrMsg= 'Total Issue Qty For All Bag Characters Do Not Match'              
      GoTo ErrHndlr              
     end              
--5717
   if @mode in ('A')              
   begin    	
    select @Row= Count('x') from inserted ins               
     where ins.TdPrtKey=@wCurrPrtn and ins.TdSrNo= 0 and ins.TdBDc= 'C'               
      and exists (Select 'x' from Txnd Bc               
      join Txnd Tc on Tc.TdPrtKey=Bc.TdPrtKey and Tc.TdTIdNo=Bc.TdTIdNo              
       and Tc.TdCoCd= Bc.TdCoCd and Tc.TdTc= Bc.TdTc and Tc.TdYy= Bc.TdYy               
       and Tc.TdChr= Bc.TdChr and Tc.TdNo= Bc.TdNo and Tc.TdSr= Bc.TdSr               
       and Tc.TdSrNo> 0               
      where ins.TdPrtKey=Bc.TdPrtKey and ins.TdTIdNo=Bc.TdTIdNo              
       and ins.TdCoCd= Bc.TdCoCd and ins.TdTc= Bc.TdTc and ins.TdYy= Bc.TdYy               
       and ins.TdChr= Bc.TdChr and ins.TdNo= Bc.TdNo and Bc.TdSrNo= 0               
       and Bc.TdRefKey< ins.TdRefKey               
      group By Bc.TdRefKey, Tc.TdRmCd, convert (decimal (12, 4), Tc.TdRmSz),               
		convert (decimal (12, 4), Tc.TdRmSz2),               convert (decimal (12, 4), Tc.TdRmSz3),               
       convert (decimal (12, 2), Tc.TdRmStkRt), Tc.TdLotNo               
      having convert (decimal (16, 3), sum(case               
        when Tc.TdFrRmLoc= 'W' and Tc.TdFrRmDc= 'C' then -Tc.TdRmWt               
        when Tc.TdFrRmLoc= 'W' and Tc.TdFrRmDc= 'D' then Tc.TdRmWt               
        when Tc.TdToRmLoc= 'W' and Tc.TdToRmDc= 'C' then -Tc.TdRmWt               
        when Tc.TdToRmLoc= 'W' and Tc.TdToRmDc= 'D' then Tc.TdRmWt               
        else 0 end)) <> convert (decimal (16, 3), 0)               
       or               
       sum(case when Tc.TdFrRmLoc= 'W' and Tc.TdFrRmDc= 'C' then -Tc.TdRmQty               
        when Tc.TdFrRmLoc= 'W' and Tc.TdFrRmDc= 'D' then Tc.TdRmQty               
        when Tc.TdToRmLoc= 'W' and Tc.TdToRmDc= 'C' then -Tc.TdRmQty               
        when Tc.TdToRmLoc= 'W' and Tc.TdToRmDc= 'D' then Tc.TdRmQty               
        else 0 end) <> 0               
      )               
    If @Row > 0              
       begin              
     Select @ErrMsg= 'Complete The RM Matching Of Previous Parent-Child Records Before Starting Another Set'              
     GoTo ErrHndlr              
    end                
   end              
	--5717
	   select @Row= Count('x') from inserted where TdPRtKEy=@wCurrPrtn and TdSrNo= 0 and TdRefYy<> ''               
       If @Row > 0              
        begin              
    Select @ErrMsg= 'RefYy Should Be Blank'              
      GoTo ErrHndlr              
     end 
	--5717
	  select @Row= Count('x') from inserted where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and TdRefKey= 0               
       If @Row > 0              
        begin              
    Select @ErrMsg= 'RefKey Cannot Be Blank'              
      GoTo ErrHndlr              
     end   
	 --5717
   If @Mode = 'C'               
   Begin                  
    select @Row=Count('x') from Deleted Del where Del.TdPrtKey=@wCurrPrtn and              
      Not Exists (Select 'x' from Inserted Ins where INs.TdPrtKey=Del.TdPrtKey               
       and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd= Del.TdCoCd               
       and Ins.TdTc= Del.TdTc and Ins.Tdyy= Del.TdYy               
       and Ins.TdChr= Del.TdChr and Ins.TdNo= Del.TdNo               
       and Ins.TdSr= Del.TdSr and Ins.TdSrNo= Del.TdSrNo               
       and Ins.TdRefKey= Del.TdRefKey)           
    If @Row > 0              
    begin              
     Select @ErrMsg='Cannot Change Ref Key'              
     GoTo ErrHndlr              
    end              
   End    
   --5717
           select @errmsg= TdBQty from inserted where TdPrtKey=@wCurrPrtn and TdSrNo= 0 and TdBDc= 'D' and              
    Not Exists (Select 'x' from Bag where BPrtKey=TdPrtKey and BIdNo=TdBIdNo and BCoCd= TdCoCd and BYy= TdBYy and              
     BChr= TdBChr and BNo= TdBNo and              
    convert(decimal(16, 1), TdBQty)= convert(decimal(16, 1), BQty))              
    If @Row> 0              
        Begin              
    Select @ErrMsg= 'Cannot Change Quantity Of Receiving Bag'              
      GoTo ErrHndlr              
        end   

    			
	Select @Row= Count('x') From Inserted Ins Where TdSrNo= 0 and              
    Exists (Select 'x' from Param where PTyp= 'BCHR' and PMCd= TdBChr and PSCd= '' and              
     PValue= 'P') and TdBDc= 'C'              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Primary Bag Will Always Be The Receiving Bag'              
    GoTo ErrHndlr              
   end              
                  
   Select @Row= Count('x') From Inserted Where TdSrNo= 0 and TdBDc= 'C' and TdBQty< 0              
     If @Row> 0              
   begin              
    Select @ErrMsg= 'Flute Bag Quantity Should Be >= 0 '              
    GoTo ErrHndlr              
   end   

   select @Row= Count('x') from               
    (Select sum(case when Tc.TdBDc= 'C' then Tc.TdBQty else 0 end) as qCrQty,               
      sum(case when Tc.TdBDc= 'D' then Tc.TdBQty else 0 end) as qDrQty               
     from Txnd Tc where exists(Select 'x' from inserted ins               
      where ins.TdBDc= 'D' and ins.TdSrNo= 0               
       and ins.TdTIdNo=Tc.TdTIdNo              
			and ins.TdRefYy=0 and ins.TdRefKey= Tc.TdRefKey)         
      and Tc.TdSrNo= 0               
     group By Tc.TdBChr) Tbl               
     having convert( decimal (16, 1), sum(qDrQty))> convert( decimal (16, 1), max(qCrQty))               

   If @Row > 0               
      begin              
    Select @ErrMsg= 'Receipt Quantity Cannot Be More Than Issue'              
    GoTo ErrHndlr              
   end              
              
	   select @Row= Count('x') from inserted Ins where TdPrtKEy=@wCurrPrtn and              
    TdSrNo= 0 and exists               
    (Select 'x' from Txnd Td where Td.TdPrtKEy=Ins.TdPrtKey and Td.TdTIdNo=Ins.TdTIdNo              
     and Td.TdCoCd= Ins.TdCoCd and Td.TdTc= Ins.TdTc               
     and Td.TdYy= Ins.TdYy and Td.TdChr= Ins.TdChr and Td.TdNo= Ins.TdNo               
     and Td.TdSrNo= 0 and Td.TdSr< Ins.TdSr and Td.TdRefKey> Ins.TdRefKey)               
   If @Row > 0               
      begin              
    Select @ErrMsg= 'RefKey Should Be >= That Of Previous Sr'              
    GoTo ErrHndlr              
   end              
              
			  
    /*** (Jen 2.13) ***/              
    /****** Cannot Delete 'BM', TdDc= 'C', TdSrNo= 0               
    If exists TdBDc= 'D', TdSrNo= 0, with same TdRefKey in same voucher ********/               
    /* Current Partition */              
    Select @Row= Count('x') from Deleted Del where 
     Exists (Select 'x' from Txnd Td where 
      Td.TdTIdNo=Del.TdTIdNo and Td.TdCoCd= Del.TdCoCd and Td.TdBDc= 'D' and Td.TdSrNo= 0 and               Td.TdRefYy=0 and
      Td.TdRefKey= Del.TdRefKey               
      /* and Td.TdSr> Del.TdSr */              
             )               
           and TdBDc= 'C' and TdSrNo= 0               

    If @Row > 0              
         Begin              
		    Select @ErrMsg= 'Cannot Delete Parent, Child Exists'              
			GoTo ErrHndlr              
         End              


--   End              
--  End
 End              
End              
 IF @wEntTyp='TxnRm'              
 BEGIN              
  --5717
  Select @Row=Count('x') From Inserted Ins where Ins.TdPrtKey=@wCurrPrtn and               
   TdBYy<> '' or TdBChr<>'' or TdBNo<>0 or TdBDc <> ''              
        or TdFrBLoc <> '' or TdToBLoc <> '' or TdBGrWt <> 0              
        or TdBQty <> 0 or TdPtQty <> 0              
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Bag related fields should be empty'              
   GoTo ErrHndlr              
  END     

  Select @Row=Count('x') From Inserted Ins              
   Where Not Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
    LocTyp<> 'P' and LocCd = Ins.TdFrRmLoc And LocValidYN='Y')              
   /* AND TdFrRmDc = 'C' */              
      If @Row>0              
  BEGIN              
              
	Select @ErrMsg='Invalid From Rm Loc'              
     GoTo ErrHndlr              
  END              
  --5717
    SELECT @Row= COUNT('x') FROM Inserted Ins WHERE TdPrtKey=@wCurrPrtn and              
   (Ins.TdFrRmLoc <> ISNULL((SELECT Del.TdFrRmLoc FROM Deleted Del              
    WHERE Ins.TdPrtKey=Del.TdPrtKey and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   OR              
   Ins.TdFrRmDc <> ISNULL((SELECT Del.TdFrRmDc FROM Deleted Del              
    WHERE Ins.TdPrtKey=Del.TdPrtKey and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   OR Ins.TdFrRmLoc ='')              
   AND NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
    LocTyp<> 'P' AND LocCd = Ins.TdFrRmLoc AND LocValidYN = 'Y')              
   AND TdFrRmDc = 'D'              
/* **** Zubin 211 (Or Condition for 'TdFrRmDc' was missing) **** */              
      IF @Row>0              
  BEGIN              
              
   SELECT @ErrMsg='Invalid From Rm Loc'              
     GOTO ErrHndlr              
  END   
  --5717
    Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Param where PTyp='DC'              
    and PMCd = Ins.TdFrRmDc and PScd='')               
    If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid From Rm DC'              
     GoTo ErrHndlr              
  END    
  --5717
    Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and Not Exists (Select 'x' From RmMst where RmCd=TdRmCd and RmPrtKey=TdPrtKey and RmZ='N')              
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid Rm Code'              
     GoTo ErrHndlr              
  END    
              
/*  ####################  Manoj # Ver: 2.0.11 Patch 01 # Date: 17/08/2005  ####################  */              
  If @wTcTyp = 'PR'              
  Begin            
   /* Current Partition */              
   Select @Row=Count('x') From Inserted Ins              
    Join Loc On LocTyp = 'XP' And LocCoCd = Ins.TdCoCd              
    And LocCd = (Case When Ins.TdFrRmDc = 'C' Then Ins.TdFrRmLoc              
      When Ins.TdFrRmDc = 'D' Then Ins.TdToRmLoc End)              
   Where Not Exists (Select 'x' From RmMst where RmCd=TdRmCd and RmZ='N' And RmValidYN = 'Y')              
       If @Row>0              
   BEGIN              
    Select @ErrMsg='Rm Code Not Defined / InValid'              
      GoTo ErrHndlr              
   END              
  End              
--   the value from X is used - except for ValidYN check is not current
  --Select @Row=Count('x') From Inserted Ins              
  -- Where (Not Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
  --  LocTyp<> 'P' and LocCd = Ins.TdToRmLoc And LocValidYn='Y'              
  --  and (@wTcTyp Not In('Rec','Cnv','Bs') Or LocTyp= 'Xc')))              
  --    If @Row>0              
  --BEGIN              
  -- Select @ErrMsg='Invalid To Rm Loc'              
  --   GoTo ErrHndlr              
  --END              
  --5717
  Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
    LocTyp<> 'P' and LocCd = Ins.TdToRmLoc And LocValidYn='Y'              
    and (@wTcTyp Not In('Rec','Cnv','Bs') Or LocTyp= 'Xc')))                  
      If @Row>0              
  BEGIN              
   Select @ErrMsg='Invalid To Rm Loc'              
     GoTo ErrHndlr              
  END    
  --5717
    SELECT @Row= COUNT('x') FROM Inserted Ins WHERE Ins.TdPrtKey=@wCurrPrtn and              
   ((              
   (Ins.TdToRmLoc <> ISNULL((SELECT Del.TdToRmLoc FROM Deleted Del              
    WHERE Ins.TdPrtKey=Del.TdPrtKey and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   OR              
   Ins.TdToRmDc <> ISNULL((SELECT Del.TdToRmDc FROM Deleted Del              
    WHERE Ins.TdPrtKey=Del.TdPrtKey and Ins.TdIdNo=Del.TdIdNo and Ins.TdCoCd = Del.TdCoCd AND              
    Ins.TdTc = Del.TdTc AND Ins.TdYy = Del.TdYy AND Ins.TdChr = Del.TdChr AND              
    Ins.TdNo = Del.TdNo AND Ins.TdSr = Del.TdSr AND Ins.TdSrNo = Del.TdSrNo), '')              
   )              
   AND (NOT EXISTS (SELECT 'x' FROM Loc WHERE LocCoCd= Ins.TdCoCd AND              
    LocTyp<> 'P' AND LocCd = Ins.TdToRmLoc AND LocValidYN = 'Y'              
    AND (@wTcTyp NOT IN('Rec','Cnv','Bs') OR LocTyp= 'Xc'))))              
   Or Ins.TdToRmLoc = '')              
   AND TdToRmDc = 'D'              
      IF @Row>0              
  BEGIN              
   SELECT @ErrMsg='Invalid To Rm Loc'              
     GOTO ErrHndlr              
  END         
  --5717
    select @row= count('x') from inserted Ins where Ins.TdPrtKey=@wCurrPrtn and /*TdSrNo<>0 and */              
    Ins.TdFrRmLoc = Ins.TdToRmLoc and (Ins.TdFrRmLoc <> '' or              
    Ins.TdToRmLoc <> '')              
    if @row>0                begin              
      select @errmsg= 'To Rm Loc and From Rm Loc cannot be same'              
      goto ErrHndlr              
    end   
	--5717
	  Select @Row=Count('x') From Inserted Ins              
   Where TdPrtKey=@wCurrPrtn and Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and              
     LocTyp= 'W' and LocCd = Ins.TdFrRmLoc)              
           or Exists (Select 'x' From Loc where LocCoCd= Ins.TdCoCd and LocTyp= 'W' and LocCd = Ins.TdToRmLoc)              
    If @Row>0              
  begin              
   Select @ErrMsg= 'From Or To Rm Loc Should Not Be ''W'' '              
   GoTo ErrHndlr              
  end 
  --5717
    Select @Row= Count('x') From Inserted Ins where Ins.TdPrtKey=@wCurrPrtn and              
   TdRjDc<> '' or TdRjQty<>0 or TdRjPtQty<> 0 or TdRjCd<> ''              
        or TdWrk<> '' or TdPrdYN<> '' or TdByLoc<> ''              
   or TdByWrk<> ''              
      If @Row> 0              
  begin              
   Select @ErrMsg= 'Rej, Prd Related Fields Should Be Empty'              
     GoTo ErrHndlr              
  end 
  --5717
     select @Row= Count('x') from Inserted Ins where Ins.TdPrtKey=@wCurrPRtn and @wTc='PR' AND @WTc='SA'      
  and Not Exists(Select * from param Where Ptyp='LABQW' and PMCd=Ins.TdRtByQW)      
   If @Row> 0              
   begin              
    Select @ErrMsg= 'Invalid Rm Qw Option'              
    GoTo ErrHndlr              
   end 
   --5717
      select @Row= Count('x') from Inserted Ins where Ins.TdPrtKey=@wCurrPRtn and @wTc='PR'       
  and Ins.TdRtByQW='Q' and (TdRmCtg<>'X' and TdRmCtg<>'D' and TdRmCtg<>'C')   
   If @Row> 0              
   begin              
    Select @ErrMsg= '"Q" Allowed in Rate By QW Option Only For Dia/CS/Accessories'              
    GoTo ErrHndlr              
   end         


END              
              
              

              
 
              
              
              
              
            
 if @mode in ('C','D') /*and  @AckYN='Y'            */
 begin              
  /* Current Partition */      
              
  update SYyMm set              
  SRmCrQty= SRmCrQty- isnull((select isnull((Del.TdRmQty), 0) from Deleted Del              
     where Del.TdFrRmDc='C' and Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd              
   and Del.TdLotNo= SLotNo and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))= SYyMm and Del.TdAck='Y' ),0),              
  SRmCrWt= SRmCrWt-isnull((select isnull((Del.TdRmWt),0) from Deleted Del              
   where Del.TdFrRmDc='C' and Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd              
     and Del.TdLotNo= SLotNo and Del.TDRmSz= SRmSz and Del.TDRmSz2= SRmSz2 and Del.TDRmSz3= SRmSz3 and Del.TdRmStkRt= SRmStkRt and Del.TdFrRmLoc=SLoc              
     and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck='Y'),0),              
  SRmDrQty= SRmDrQty-isnull((select isnull((Del.TdRmQty),0) from Deleted Del              
    where Del.TdFrRmDc='D' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd              
      and Del.TdLotNo= SLotNo and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdFrRmLoc=SLoc              
      and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck='Y'),0),              
  SRmDrWt= SRmDrWt-isnull((select isnull((Del.TdRmWt),0) from Deleted Del              
      where Del.TdFrRmDc='D' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd              
      and Del.TdLotNo= SLotNo and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck = 'Y' ),0) 
  from Deleted       
   where Deleted.TdCoCd= SCoCd and Deleted.TdRmCd= SRmCd and Deleted.TdLotNo= SLotNo              
   and Deleted.TDRmSz= SRmSz and Deleted.TDRmSz2= SRmSz2 and Deleted.TDRmSz3= SRmSz3 and Deleted.TdRmStkRt= SRmStkRt and Deleted.TdFrRmLoc= SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm              
   and (Deleted.TdFrRmDc='D'  or Deleted.TdFrRmDc='C') and       
   ((@Mode='C' and Exists(Select 'x' From Inserted ins where ins.TdCoCd =Deleted.TdCoCd and ins.TdTc=Deleted.TdTc and ins.TdYy=Deleted.TdYy and      
         ins.TdChr=Deleted.TdChr and ins.TdNo=Deleted.TdNo and ins.TdSr=Deleted.TdSr and ins.TdSrNo=Deleted.TdSrNo and          
      (ins.TdCoCd<>Deleted.TdCoCd Or ins.TdRmCd<>Deleted.TdRmCd Or ins.TdLotNo<>Deleted.TdLotNo      
        Or  ins.TDRmSz<>Deleted.TDRmSz or ins.TDRmSz2<>Deleted.TDRmSz2 or ins.TDRmSz3<>Deleted.TDRmSz3 or ins.TdRmStkRt<>Deleted.TdRmStkRt Or ins.TdFrRmLoc<>Deleted.TdFrRmLoc Or       
         Ins.TdRmWt<>Deleted.TdRmWt Or Ins.TdRmQty<>Deleted.TdRmQty Or  Ins.TdFrRmDC<>Deleted.TdFrRmDC Or  
         Ins.TdAck <> Deleted.TdAck or  
         convert(int,convert(varchar(6),ins.TdDt,112))<>convert(int,convert(varchar(6),Deleted.TdDt,112))))) Or @Mode='D')      
      
       
 end              

 if @mode in ('A','C')  /* and @AckYN= 'Y'           */
 begin              
  /* Current Partition */              
  insert into SYyMm(SPrtKey,SCoCd,SRmCd,SLotNo,SRmSz,SRmSz2, SRmSz3, SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)              
   select distinct TdPrtKey, TdCoCd, TdRmCd,TdLotNo,TdRmSz,TdRmSz2, TdRmSz3, TdRmStkRt,TdFrRmLoc,convert(int,convert(varchar(6),TdDt,112)),0,0,0,0              
   from inserted       
   where not exists(select 'x' from  SYyMm where SCoCd= TdCoCd and SRmCd=TdRmCd and SLotNo= TdLotNo              
   and SRmSZ=TdRmSz and SRmSZ2=TdRmSz2 and SRmSZ3=TdRmSz3 and SRmStkRt=TdRmStkRt and SLoc=TdFrRmLoc and              
   SYyMm= convert(int,convert(varchar(6),TdDt,112))) and            TdAck='Y' and  
   ((@wEntTyp= 'TxnBag' and TdSrNo<> 0) or (@wEntTyp= 'TxnRm' and TdSrNo= 0))
   
              
  /* Current Partition */               
  update SYyMm set              
  SRmCrQty=SRmCrQty+isnull((select isnull((Ins.TdRmQty),0) from Inserted Ins       
   where Ins.TdFrRmDc='C' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdFrRmLoc=SLoc              
  and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck='Y'),0),              
  SRmCrWt=SRmCrWt+isnull((select isnull((Ins.TdRmWt),0) from Inserted Ins       
   where Ins.TdFrRmDc='C' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck='Y'),0),              
  SRmDrQty=SRmDrQty+isnull((select isnull((Ins.TdRmQty),0) from Inserted Ins       
   where Ins.TdFrRmDc='D'  and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck='Y'),0),              
  SRmDrWt=SRmDrWt+isnull((select isnull((Ins.TdRmWt),0) from Inserted Ins       
   where Ins.TdFrRmDc='D'  and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdFrRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck='Y'),0)              
  from Inserted       
   where Inserted.TdCoCd= SCoCd and Inserted.TdRmCd= SRmCd and Inserted.TdLotNo=SLotNo              
   and Inserted.TDRmSz= SRmSz and Inserted.TDRmSz2= SRmSz2 and Inserted.TDRmSz3= SRmSz3  and Inserted.TdRmStkRt= SRmStkRt              
   and Inserted.TdFrRmLoc= SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm              
   and (Inserted.TdFrRmDc='D' or Inserted.TdFrRmDc='C')  and
   ((@Mode='C' and Exists(Select 'x' From Deleted Del where Inserted.TdCoCd =Del.TdCoCd and Inserted.TdTc=Del.TdTc and Inserted.TdYy=Del.TdYy and      
         Inserted.TdChr=Del.TdChr and Inserted.TdNo=Del.TdNo and Inserted.TdSr=Del.TdSr and Inserted.TdSrNo=Del.TdSrNo  and                   
      (Inserted.TdCoCd<>Del.TdCoCd Or Inserted.TdRmCd<>Del.TdRmCd Or Inserted.TdLotNo<>Del.TdLotNo      
        Or  Inserted.TDRmSz<>Del.TDRmSz or Inserted.TDRmSz2<>Del.TDRmSz2 or Inserted.TDRmSz3<>Del.TDRmSz3 or Inserted.TdRmStkRt<>Del.TdRmStkRt Or Inserted.TdFrRmLoc<>Del.TdFrRmLoc  or      
        Inserted.TdRmWt<>Del.TdRmWt Or Inserted.TdRmQty<>Del.TdRmQty Or Inserted.TdFrRmDC<>Del.TdFrRmDC Or 
          Inserted.TdAck <> Del.TdAck or      
         convert(int,convert(varchar(6),Inserted.TdDt,112))<>convert(int,convert(varchar(6),Del.TdDt,112))))) Or @Mode='A')             

end              
              
 /*** Update SYyMm from TdToRmLoc ***/              
 --select @AckYN= Del.TdAck  from deleted Del 	
 if @mode in ('C','D')  /*and @AckYN= 'Y'            */
 begin              
  /* Current Partition */              
  update SYyMm set              
  SRmCrQty=SRmCrQty-isnull((select isnull(sum(Del.TdRmQty),0) from Deleted Del       
   where Del.TdToRmDc='C' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd and Del.TdLotNo=SLotNo              
   and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and  Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck='Y'),0)              ,
  SRmCrWt=SRmCrWt-isnull((select isnull(sum(Del.TdRmWt),0) from Deleted Del       
   where Del.TdToRmDc='C' and Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd and Del.TdLotNo=SLotNo              
   and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3   and Del.TdRmStkRt=SRmStkRt and Del.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck='Y'),0)              ,
  SRmDrQty=SRmDrQty-isnull((select isnull(sum(Del.TdRmQty),0) from Deleted Del       
   where Del.TdToRmDc='D' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd and Del.TdLotNo=SLotNo              
   and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck = 'Y'),0)    ,              
  SRmDrWt=SRmDrWt-isnull((select isnull(sum(Del.TdRmWt),0) from Deleted Del       
   where Del.TdToRmDc='D' and Del.TdCoCd= SCoCd and Del.TdRmCd=SRmCd and Del.TdLotNo=SLotNo              
   and Del.TDRmSz=SRmSz and Del.TDRmSz2=SRmSz2 and Del.TDRmSz3=SRmSz3 and Del.TdRmStkRt=SRmStkRt and Del.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Del.TdAck = 'Y')  ,0)            
  from Deleted       
   where Deleted.TdCoCd= SCoCd and Deleted.TdRmCd= SRmCd and Deleted.TdLotNo=SLotNo              
   and Deleted.TDRmSz= SRmSz and Deleted.TDRmSz2= SRmSz2 and Deleted.TDRmSz3= SRmSz3 and Deleted.TdRmStkRt= SRmStkRt and Deleted.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),Tddt,112))= SYyMm              
   and (Deleted.TdToRmDc='D' or Deleted.TdToRmDc='C')  and      
   ((@Mode='C' and Exists(Select 'x' From Inserted ins where ins.TdCoCd =Deleted.TdCoCd and ins.TdTc=Deleted.TdTc and ins.TdYy=Deleted.TdYy and      
         ins.TdChr=Deleted.TdChr and ins.TdNo=Deleted.TdNo and ins.TdSr=Deleted.TdSr and ins.TdSrNo=Deleted.TdSrNo and          
      (ins.TdCoCd<>Deleted.TdCoCd Or ins.TdRmCd<>Deleted.TdRmCd Or ins.TdLotNo<>Deleted.TdLotNo      
        Or  ins.TDRmSz<>Deleted.TDRmSz or ins.TDRmSz2<>Deleted.TDRmSz2 or ins.TDRmSz3<>Deleted.TDRmSz3 or    ins.TdRmStkRt<>Deleted.TdRmStkRt Or ins.TdToRmLoc<>Deleted.TdToRmLoc Or       
        Ins.TdRmWt<>Deleted.TdRmWt Or Ins.TdRmQty<>Deleted.TdRmQty Or Ins.TdToRmDC<>Deleted.TdToRmDC Or   Ins.TdAck <> Deleted.TdAck or      
         convert(int,convert(varchar(6),ins.TdDt,112))<>convert(int,convert(varchar(6),Deleted.TdDt,112))))) Or @Mode='D')      
      
             
end              
              
 --select @AckYN= Ins.TdAck  from inserted Ins 	             
 if @mode in ('A','C') /*and @AckYN = 'Y'*/
 begin              
  /* Current Partition */              
  insert into SYyMm(SPrtKey,SCoCd, SRmCd,SLotNo,SRmSz,SRmSz2, SRmSz3, SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)              
  select distinct TdPrtKey,TdCoCd, TdRmCd,TdLotNo,TdRmSz,TdRmSz2, TdRmSz3, TdRmStkRt, TdToRmLoc, convert(int,convert(varchar(6),TdDt,112)), 0, 0, 0, 0              
   from inserted       
   where not exists(select 'x' from  SYyMm where SCoCd= TdCoCd and SRmCd=TdRmCd and SLotNo=TdLotNo              
   and SRmSZ=TdRmSz and SRmStkRt=TdRmStkRt and SLoc=TdToRmLoc and    SRmSZ2=TdRmSz2 and    SRmSZ3=TdRmSz3        and       
   SYyMm=convert(int,convert(varchar(6),TdDt,112))) and TdAck = 'Y' and             
   ((@wEntTyp= 'TxnBag' and TdSrNo<> 0) or (@wEntTyp= 'TxnRm' and TdSrNo= 0))              
    
   
  /* Current Partition */              
  update SYyMm set              
  SRmCrQty=SRmCrQty+isnull((select isnull((Ins.TdRmQty),0) from Inserted Ins       
   where Ins.TdToRmDc='C' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck = 'Y'),0),              
  SRmCrWt=SRmCrWt+isnull((select isnull((Ins.TdRmWt),0) from Inserted Ins       
   where Ins.TdToRmDc= 'C' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck = 'Y'),0) ,             
  SRmDrQty=SRmDrQty+isnull((select isnull((Ins.TdRmQty),0) from Inserted Ins       
   where Ins.TdToRmDc='D' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck = 'Y'),0),              
  SRmDrWt=SRmDrWt+isnull((select isnull((Ins.TdRmWt),0) from Inserted Ins       
   where Ins.TdToRmDc='D' and Ins.TdCoCd= SCoCd and Ins.TdRmCd=SRmCd and Ins.TdLotNo=SLotNo              
   and Ins.TDRmSz=SRmSz and Ins.TDRmSz2=SRmSz2 and Ins.TDRmSz3=SRmSz3 and Ins.TdRmStkRt=SRmStkRt and Ins.TdToRmLoc=SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm and Ins.TdAck = 'Y'),0)              
  from Inserted       
   where Inserted.TdCoCd= SCoCd and Inserted.TdRmCd=SRmCd and Inserted.TdLotNo=SLotNo              
   and Inserted.TDRmSz=SRmSz and Inserted.TDRmSz2=SRmSz2 and Inserted.TDRmSz3=SRmSz3 and Inserted.TdRmStkRt=SRmStkRt and Inserted.TdToRmLoc= SLoc              
   and convert(int,convert(varchar(6),TdDt,112))=SYyMm              
   and (Inserted.TdToRmDc='D' or Inserted.TdToRmDc='C')  and      
 ((@Mode='C' and Exists(Select 'x' From Deleted Del where Inserted.TdCoCd =Del.TdCoCd and Inserted.TdTc=Del.TdTc and Inserted.TdYy=Del.TdYy and      
         Inserted.TdChr=Del.TdChr and Inserted.TdNo=Del.TdNo and Inserted.TdSr=Del.TdSr and Inserted.TdSrNo=Del.TdSrNo and                   
      (Inserted.TdCoCd<>Del.TdCoCd Or Inserted.TdRmCd<>Del.TdRmCd Or Inserted.TdLotNo<>Del.TdLotNo      
        Or  Inserted.TDRmSz<>Del.TDRmSz Or  Inserted.TDRmSz2<>Del.TDRmSz2 or  Inserted.TDRmSz3<>Del.TDRmSz3 or Inserted.TdRmStkRt<>Del.TdRmStkRt Or Inserted.TdToRmLoc<>Del.TdToRmLoc Or       
          Inserted.TdRmWt<>Del.TdRmWt Or Inserted.TdRmQty<>Del.TdRmQty Or  Inserted.TdToRmDC<>Del.TdToRmDC Or 
		  Inserted.TdAck <> Del.TdAck or  
         convert(int,convert(varchar(6),Inserted.TdDt,112))<>convert(int,convert(varchar(6),Del.TdDt,112))))) Or @Mode='A')                   
      
 end              
              
 /* urmi NegStk Start*/              
 IF @wEntTyp in ('TxnRm', 'TxnBag')              
 BEGIN
 
 if @mode in ('A','C')              
 begin              
 
  Select @Row=Count('x')               
     From Inserted Ins               
      Join Txn on TIdNo=Ins.TdTIdNo and TCoCd= Ins.TdCoCd and TTc= Ins.TdTc and TYy= ins.TdYy and              
        TChr= Ins.TdChr and TNo= Ins.TdNo               
      Join Head On HCoCd=TCoCd And HCd= @wHCd And HNegStkYn= 'N'              
      Join RmMst On RmCd=Ins.TdRmCd And               
        (Case When hNegStkRmCtg = '' Then ','+RmCtg+','               
        Else (Select ','+PDesc225+',' From Param               
            Where PTyp='GLDAS' And PMCd=hNegStkRmCtg) End) Like '%,'+RmCtg+',%'              
  Where Exists (Select 'X' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and              
        HNegStkYn= 'N')              
  and not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and               
    PMCd= Ins.ModUsr and UaMnuCd = Ins.TdTc and UaSeoYn= 'Y' and PValue= TSeoPwd)               
  and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and TSeoPwd= PValue1)               
  and Exists (Select 'X' From Loc where LocCoCd= Ins.TdCoCd and LocTyp = 'S'              
   and LocCd = (case when Ins.TdFrRmDC = 'C' then Ins.TdFrRmLoc              
                         when Ins.TdFrRmDC = 'D' then Ins.TdToRmLoc              
          else '' end) )              
  and ((Select CAST(SUM(SRmDrWt)-SUM(SRmCrWt) As Decimal(10,3)) from SYyMm              
    where Ins.TdCoCd= SCoCd and Ins.TdRmCd= SRmCd and Ins.TdLotNo= SLotNo and              
     Ins.TDRmSz= SRmSz and              Ins.TDRmSz2= SRmSz2 and              Ins.TDRmSz3= SRmSz3 and              
     (case when Ins.TdFrRmDC = 'C' then Ins.TdFrRmLoc              
 when Ins.TdFrRmDC = 'D' then Ins.TdToRmLoc              
     else '' end )= SLoc              
       )
      ) < 0              
  And (Exists (Select 'X' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and              
       HAvgSRtYN = 'Y')              
       Or              
       Exists (Select 'X' from Param Join RmMst On PTyp = 'RMSCTG' AND RmCd = Ins.TdRmCd And              
   PMCD = RMCTG And PSCd = RMSCTG where PValue1 = 'N'))              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Only Authorized Users can enter Negative Stock For This Rm Ctg'              
     GoTo ErrHndlr              
       end              
              
  Select @Row=Count('x') From Inserted Ins               
    join Txn on TIdNo=Ins.TdTIdNo and TCoCd= Ins.TdCoCd and              
       TTc= Ins.TdTc and TYy= ins.TdYy and TChr= Ins.TdChr and TNo= Ins.TdNo              
      Join Head On HCoCd=TCoCd And HCd= @wHCd And HNegStkYn= 'N'              
      Join RmMst On RmCd=Ins.TdRmCd And               
        (Case When hNegStkRmCtg = '' Then ','+RmCtg+','               
    Else (Select ','+PDesc225+',' From Param               
            Where PTyp='GLDAS' And PMCd=hNegStkRmCtg) End) Like '%,'+RmCtg+',%'              
  Where Exists (Select 'X' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and              
        HNegStkYn= 'N')              
  and not Exists (select 'x' from Param Join UsrAccess Ua on Ua.UaUsrCd = PMCd where PTyp= 'USR' and               
    PMCd= ins.ModUsr and Ua.UaMnuCd = ins.TdTc and Ua.UaSeoYn= 'Y' and PValue= TSeoPwd)               
  and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and TSeoPwd= PValue1)               
  and Exists (Select 'X' From Loc where LocCoCd= Ins.TdCoCd and LocTyp = 'S'              
   and LocCd = (case when Ins.TdFrRmDC = 'C' then Ins.TdFrRmLoc              
                         when Ins.TdFrRmDC = 'D' then Ins.TdToRmLoc              
          else '' end) )              
              
  and ( (Select CAST(SUM(SRmDrWt)-SUM(SRmCrWt) As Decimal(10,3)) from SYyMm              
    where Ins.TdCoCd= SCoCd and Ins.TdRmCd= SRmCd and Ins.TdLotNo= SLotNo and              
     Ins.TDRmSz= SRmSz and Ins.TDRmSz2= SRmSz2 and Ins.TDRmSz3= SRmSz3 and Ins.TdRmStkRt =SRmStkRt and              
     (case when Ins.TdFrRmDC = 'C' then Ins.TdFrRmLoc              
        when Ins.TdFrRmDC = 'D' then Ins.TdToRmLoc              
      else '' end )= SLoc              
     )/*- Ins.TdRmWt */              
      ) < 0              
  And Exists (Select 'X' From Head where HCoCd= Ins.TdCoCd and HCd= @wHCd and              
       HAvgSRtYN = 'N')              
  And Exists (Select 'X' from Param Join RmMst On PTyp = 'RMSCTG' AND RmCd = Ins.TdRmCd And              
   PMCD = RMCTG And PSCd = RMSCTG where PValue1 = 'Y')              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Only Authorized Users can enter Negative Stock For This Rm Ctg'              
     GoTo ErrHndlr              
       end              
              
              
      end               
      if @mode in ('D')              
      begin              
  Select @Row=Count('x') From Deleted Del join Txn on 
       TIdNo=Del.TdTIdNo and TCoCd= Del.TdCoCd and              
       TTc= Del.TdTc and TYy= Del.TdYy and              
       TChr= Del.TdChr and TNo= Del.TdNo              
      Join Head On HCoCd=TCoCd And HCd= @wHCd And HNegStkYn= 'N'              
      Join RmMst On RmCd=Del.TdRmCd And               
        (Case When hNegStkRmCtg = '' Then ','+RmCtg+','               
         Else (Select ','+PDesc225+',' From Param               
            Where PTyp='GLDAS' And PMCd=hNegStkRmCtg) End) Like '%,'+RmCtg+',%'              
  Where Exists (Select 'X' From Head where HCoCd= Del.TdCoCd and HCd= @wHCd and              
        HNegStkYn= 'N')              
  and not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and               
    PMCd= Del.ModUsr and UaMnuCd = Del.TdTc and UaSeoYn= 'Y' and PValue= TSeoPwd)               
  and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and TSeoPwd= PValue1)               
  and Exists (Select 'X' From Loc where LocCoCd= Del.TdCoCd and LocTyp = 'S'              
   and LocCd = (case when Del.TdFrRmDC = 'D' then Del.TdFrRmLoc              
                          when Del.TdFrRmDC = 'C' then Del.TdToRmLoc              
          else '' end) )              
  and ( (Select CAST(SUM(SRmDrWt)-SUM(SRmCrWt) As Decimal(10,3)) from SYyMm              
    where Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd and Del.TdLotNo= SLotNo and              
     Del.TDRmSz= SRmSz and              Del.TDRmSz2= SRmSz2 and              Del.TDRmSz3= SRmSz3 and              
     (case when Del.TdFrRmDC = 'D' then Del.TdFrRmLoc              
     when Del.TdFrRmDC = 'C' then Del.TdToRmLoc              
     else '' end )= SLoc              
       ) /*- Del.TdRmWt */              
      ) < 0              
  And (Exists (Select 'X' From Head where HCoCd= Del.TdCoCd and HCd= @wHCd and              
       HAvgSRtYN = 'Y')              
    Or              
       Exists (Select 'X' from Param Join RmMst On PTyp = 'RMSCTG' AND RmCd = Del.TdRmCd And              
   PMCD = RMCTG And PSCd = RMSCTG where PValue1 = 'N'))              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Only Authorized Users can enter Negative Stock For This Rm Ctg'              
     GoTo ErrHndlr              
  end              
              
  Select @Row=Count('x') From Deleted Del join Txn on TIdNo=Del.TdTIdNo              
       and TCoCd= Del.TdCoCd and TTc= Del.TdTc and TYy= Del.TdYy and              
       TChr= Del.TdChr and TNo= Del.TdNo              
      Join Head On HCoCd=TCoCd And HCd= @wHCd And HNegStkYn= 'N'              
      Join RmMst On RmCd=Del.TdRmCd And               
        (Case When hNegStkRmCtg = '' Then ','+RmCtg+','               
          Else (Select ','+PDesc225+',' From Param             
            Where PTyp='GLDAS' And PMCd=hNegStkRmCtg) End) Like '%,'+RmCtg+',%'              
  Where Exists (Select 'X' From Head where HCoCd= Del.TdCoCd and HCd= @wHCd and              
        HNegStkYn= 'N')              
  and not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and               
    PMCd= Del.ModUsr and UaMnuCd = Del.TdTc and UaSeoYn= 'Y' and PValue= TSeoPwd)               
  and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and TSeoPwd= PValue1)               
  and Exists (Select 'X' From Loc where LocCoCd= Del.TdCoCd and LocTyp = 'S'              
   and LocCd = (case when Del.TdFrRmDC = 'D' then Del.TdFrRmLoc              
                          when Del.TdFrRmDC = 'C' then Del.TdToRmLoc              
          else '' end) )              
              
  and ((Select CAST(SUM(SRmDrWt)-SUM(SRmCrWt) As Decimal(10,3)) from SYyMm              
    where Del.TdCoCd= SCoCd and Del.TdRmCd= SRmCd and Del.TdLotNo= SLotNo and              
     Del.TDRmSz = SRmSz and Del.TDRmSz2 = SRmSz2 and Del.TDRmSz3 = SRmSz3 and Del.TdRmStkRt = SRmStkRt and              
     (case when Del.TdFrRmDC = 'D' then Del.TdFrRmLoc              
     when Del.TdFrRmDC = 'C' then Del.TdToRmLoc              
     else '' end )= SLoc              
       ) /*- Del.TdRmWt */              
      ) < 0              
  And Exists (Select 'X' From Head where HCoCd= Del.TdCoCd and HCd= @wHCd and              
       HAvgSRtYN = 'N')              
  And Exists (Select 'X' from Param Join RmMst On PTyp = 'RMSCTG' and RmCd = Del.TdRmCd And              
       PMCD = RMCTG And PSCd = RMSCTG where PValue1 = 'Y')              
  If @Row> 0              
  begin              
   Select @ErrMsg= 'Only Authorized Users can enter Negative Stock For This Rm Ctg'              
     GoTo ErrHndlr              
  End              
      End  
	  END              
 

 /**** urmi NegStk end ****/              
              
 If @wEntTyp = 'TxnBag' And @wTcTyp= 'BM'               
 Begin              
  If @Mode in ('A', 'C')              
  Begin              
   Update Bag               
    Set BCls = 'Y',              
                      BClsDt= (Select TDt From Txn Where 
          TIdNo=ins.TdTIdNo and TCoCd= ins.TdCoCd And TTc = ins.TdTc               
          And TYy = ins.TdYy And TChr= ins.TdChr And TNo= ins.TdNo)               
   From inserted ins              
   Where BIdNo=ins.TdBIdNo and BCoCd= ins.TdCoCd And BYy= ins.TdBYy And BChr= ins.TdBChr               
    And BNo= ins.TdBNo And ins.TdSrNo= 0               
    And Convert( Decimal (16, 1), BQty) = Convert( Decimal (16, 1), 0)               
    And Convert( Decimal (16, 3), BGrWt) = Convert( Decimal (16, 3), 0)               
    And BPtQty = 0               
    And (Select PValue From Param Where PTyp= 'BCHR' And PMCd= ins.TdBChr And PSCd= '') in ('C', 'F')               
  End              
 End  
  If @wEntTyp= 'TxnBag' and @wTcTyp= 'MLT'              
  Begin              
   If @Mode= 'D'              
   Begin              
    Update Bag               
     Set BCls = 'N',              
         BClsDt = '01/01/80'              
    From Deleted Del              
    Where BIdNo=Del.TdBIdNo and BCoCd= Del.TdCoCd And BYy= Del.TdBYy And BChr= Del.TdBChr And BNo= Del.TdBNo               
--     And Convert(Decimal(16, 3), BGrWt) <> Convert(Decimal(16, 3), 0)               
   End              
  End              

 If @wEntTyp = 'TxnBag' And @wTcTyp= 'MLT'               
 Begin              

  If @Mode in ('A', 'C')              
  Begin              
   --update bag set bOpnDesc=(select Convert(Decimal (16, 3), BGrWt) from bag where bidno=(select inserted.tdbidno from inserted))
   --where BIdNo=18168

   Update Bag               
    Set BCls = 'Y',              
                      BClsDt= (Select TDt From Txn Where 
          TIdNo=ins.TdTIdNo and TCoCd= ins.TdCoCd And TTc = ins.TdTc               
          And TYy = ins.TdYy And TChr= ins.TdChr And TNo= ins.TdNo)               
   From inserted ins              
   Where BIdNo=ins.TdBIdNo and BCoCd= ins.TdCoCd And BYy= ins.TdBYy And BChr= ins.TdBChr               
    And BNo= ins.TdBNo             
    And Convert(Decimal (16, 1), BQty) = Convert( Decimal (16, 1), 0)               
    And Convert(Decimal (16, 3), BGrWt) = Convert( Decimal (16, 3), 0)               
    And BPtQty = 0               
  End              
 End  
Return              
ErrHndlr:              
 BEGIN              
  Select @ErrMsg='|'+@ErrMsg              
  RaisError(@ErrMsg,16,1)              
  Rollback Transaction              
 END 
End



GOGO
ALTER TRIGGER [dbo].[OrdDsg_INUPDL] ON [dbo].[OrdDsg]     
   FOR INSERT,UPDATE,DELETE     NOT FOR REPLICATION
As    
BEGIN    
 DECLARE @ROW INT    
 DECLARE @mode char(1)    
 DECLARE @errmsg varchar(255)    
 DECLARE @modUsr char(5)
 Declare @wSelf VarChar(8)    
 Declare @wSysDb Varchar(3)    
 DECLARE @ReplAgt BIT    
 DECLARE @wCurrPrtn VarChar(1)    
 Declare @wInFob Float  
 Declare @wInFobRs Float  
 Declare @wInLclOctroiRs Float  
 Declare @wInLclVatRs Float  
 Declare @wInLclCstRs Float  
 DECLARE @wTcTyp varchar(3)   
 declare @wIsJLTc char(1)
 IF @@ROWCOUNT=0   RETURN    
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'OrdDsg' )    
  RETURN    
    
 Select @mode= case (select count('x') from inserted) when 0 then 'D'    
        else    
    case (select count('x') from deleted) when 0 then 'A' else 'C' end    
        end     
    
 Select @wSysDb=SysDb From SysDb where SysId='CRM'     
 Select @wCurrPrtn = 'C'       
 Set @wSysDb=IsNull(@wSysDb,'')     
 declare @wModUsr varchar(5)
 select @wModUsr=''
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))    
 If @ReplAgt = 1 Return
 Select @wSelf='ZSELF'     
 if @mode in ('A','C')  
 begin  
  select @wModUsr=max(inserted.ModUsr), @wTcTyp= max(PValue),@wIsJLTc = max(PValue1) from inserted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''  
  if @wTcTyp = '' select @wTcTyp = 'N'

 end  
  
 if @mode in ('D')  
 begin  
  select @wTcTyp=max(PValue), @wIsJLTc = max(PValue1) from deleted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''   
 End  
 --5717
 /******** Only One TcTyp records can be Inserted at a time.********/   
 select @Row= Count(distinct(PValue)) from inserted Join param On  
  PTyp= 'TC' and PMCd= OdTc and PScd= '' Where OdPrtKey=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'  
  GoTo ErrHndlr  
 end  
 --5717 
 /******** Only One TcTyp records can be Deleted at a time. *********/   
 select @Row= Count(distinct(PValue)) from deleted Join param on  
  PTyp= 'TC' and PMcd= OdTc and PSCd='' Where OdPrtKey=@wCurrPrtn  
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted on'  
  GoTo ErrHndlr  
 end  
   
 Select @Row= Count('x') from deleted where OdTc = 'JRI' and exists 
 (Select 'a' from Txnd where TdCoCd = OdCoCd and TdBYy = OdDmStkYy and TdBChr = OdDmStkChr and TdBNo = OdDmStkNo and (Txnd.ModDt > deleted.ModDt or (Txnd.ModDt = deleted.ModDt and Txnd.ModTime > Deleted.ModTime)))
 If @Row > 0  
 begin   
  Select @ErrMsg= 'Cannot Modify/Delete As Bag Transactions exists for Repair Bag'  
  GoTo ErrHndlr  
 end

 if @mode in ('C','D') and @wTcTyp = 'JT'
 Begin
	 Select @row= count('x') from deleted del where 
	   not exists(Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdDmCd = del.OdDmCd 
	   and ins.OdSfx= del.OdSfx and ins.OdDmSz = del.OdDmSz and ins.OdDmStkCoCd = del.OdDmStkCoCd and ins.OdDmStkYy = del.OdDmStkYy
	   and ins.OdDmStkChr = del.OdDmStkChr and ins.OdDmStkNo= del.OdDmStkNo) and
	   exists (Select 'a' from OrdRm where OrOdMstIdNo = del.OdIdNo)
	 If @row > 0     
	 begin    
	  Select @ErrMsg= 'Cannot change DmCd, Sfx, Sz or DmStkNum or delete when this is the master record for JRO transactions'  
	  goto ErrHndlr    
	 End   
 End
     
  If @Mode ='D'     
  Begin    
   Delete OrdRm From deleted, OrdRm where OdIdNo=OrOdIdNo and OdCoCd = OrCoCd and OdTc = OrTc and     
    OdYy = OrYy and OdChr = OrChr and OdNo = OrNo and OdSr= OrSr 
     
   Delete OrdLab From deleted, OrdLab where OdIdNo=OlOdIdNo and OdCoCd = OlCoCd and OdTc = OlTc and     
    OdYy = OlYy and OdChr = OlChr and OdNo = OlNo and OdSr= OlSr 
     
   Delete OrdAmd From deleted, OrdAmd where OdIdNo=OaOdIdNo and OdCoCd = OaCoCd and OdTc = OaTc and     
    OdYy = OaYy and OdChr = OaChr and OdNo = OaNo and OdSr= OaSr 
     
   Delete OrdTgt From deleted, OrdTgt where OdCoCd = OtCoCd and OdTc = OtTc and     
    OdYy = OtYy and OdChr = OtChr and OdNo = OtNo and OdSr = OtSr 
     
   Delete OrdPln From deleted, OrdPln where OdCoCd = OpCoCd and OdTc = OpTc and     
    OdYy = OpYy and OdChr = OpChr and OdNo = OpNo and OdSr = OpSr     
    
   Delete OrdCT From deleted, OrdCT where OdIdNo=OctOdIdNo and OdCoCd = OctCoCd and OdTc = OctTc and     
    OdYy = OctYy and OdChr = OctChr and OdNo = OctNo and OdSr = OctSr 
    End      
 --End    
	  
	Select @row= count('x') from inserted ins where ins.OdHld= 'Y'     
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

	  If @mode = 'C'    and (update(OdMinDiatolwt) or update(odMaxdiatolWT))
	  Begin      
	 Select @row=count('x') from inserted Where OdMinDiaTolWt > OdMaxDiaTolWt  
	 If @row > 0  
	 Begin  
	  Select @Errmsg='Minimunm Diamond Tolerance cannot be Greater Than Maximum Diamond Tolerance'  
	  Goto ErrHndlr  
	 end   
    
	 
 End    
    
  if @mode = 'A'     
  Begin    
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo 
	
    where OmReYy= '' And OmReChr= '' And OmReNo=0 And @wIsJLTc = 'N' and
    
    not exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp     
        and DmHld= 'N'  and DmCd= OdDmCd and DmSz= ''  )    
  if @row> 0     
  begin    
   select @errmsg= 'Design/Sketch Code Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from inserted     
    where @wIsJLTc = 'N' and
    not exists (Select PValue from Param where PTyp= 'DMTCTYP' and PMCd= OdDMTcTyp)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Tc Type is Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end
    
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo 
	
   where OmReYy= '' And OmReChr= '' And OmReNo=0 And  @wIsJLTc = 'N' and    
	@wIsJLTC='N' and 
    Exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp     
     and (DmHld= 'Y' ) and DmCd= OdDmCd and DmSz= OdDmSz )    
  if @row> 0     
  begin    
   select @errmsg= 'This Design/Sketch Size Is On Hold/Invalid'    
   goto ErrHndlr    
  end  
  --5717  
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) <> 'QB' And  
  (Select PMCd from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT') <> 'sri')   
  and (OdBYy <> '' or OdBChr <> '' or OdBNo <> 0)    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Should Be Blank For This Voucher Type'    
   goto ErrHndlr    
  end 

  -- Check if Not exists part with OrdDsg makes sense
  select @row= count('x') from inserted Ins where 
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('QB'))   
 and not exists (Select 'x' From Bag Where BCoCd= OdCoCd and BYy= OdBYy     
      and BChr = OdBChr And BNo= OdBNo and BLoc = 'PFG'     
      and BOdDmCd= OdDmCd )    
  /* And Not Exists (Select 'x' From OrdDsg Od Where Od.OdCoCd=ins.OdCoCd And Od.OdTc In ('QB') And     
      Od.OdBYy=Ins.OdBYy And Od.OdBChr=Ins.OdBChr And Od.OdBNo=Ins.OdBNo And Od.OdPrtKey=Ins.OdPrtKey)  */  
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Not Defined For the Design'    
   goto ErrHndlr    
  end    
end 
 if @mode in ('A','C') and @wModUsr<>'MIC'     
 Begin    
  select @row= count('x') from inserted where not exists (select 'x' from OrdMst where     
     OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and     
     OmChr= OdChr and OmNo= OdNo )    
  if @row> 0     
  begin    
   select @errmsg= 'Voucher Not Defined In Order Master'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdSr<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Design Serial Number Cannot Be <= Zero'    
   goto ErrHndlr    
  end  
  --5717
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
    
  select @row= count('x') from inserted Ins where OdDmSz<> '' and     
   Ins.OdDmSz <> IsNull((Select Del.OdDmSz From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
    Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy     
    And Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr ), '') And     
   not exists (select 'x' from Param where PTyp= 'DMSZ' and PMCd= OdDmSz And PValidYn = 'Y')    
  if @row> 0     
  begin    
   select @errmsg= 'Design Size Not Defined / Invalid.'    
   goto ErrHndlr    
  end    

  Select @row= count('x') from inserted where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and not exists (select 'x' from vParam where vPCoCd= OdCoCd and     
     vPTyp= 'PRDSEQ' and vPMCd= OdPrdSeq )    
  if @row> 0     
  begin    
   select @errmsg= 'Production Sequence Not Defined'    
   goto ErrHndlr    
  end   
    
  select @row= count('x') from inserted where OdPrdQty<  OdFgQty     
  if @row> 0     
  begin    select @errmsg= 'Production Qty Should Be >= Finished Goods Qty'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and    
   not exists (select 'x' from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd )    
  if @row> 0     
  begin    select @errmsg= 'Priority Code Not Defined'    
   goto ErrHndlr    
  end   
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPRtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdHld)    
  if @row> 0     
  begin    
   select @errmsg= 'Enter On Hold (Y/N)'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where OdPRtKey=@wCurrPRtn and OdHld= 'N' and OdHldDesc<> ''     
  if @row> 0     
  begin    
   select @errmsg= 'Holding Desc Should Be Entered Only When Design Is On Hold'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdMulBy<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Multiplying Factor Cannot Be <= Zero'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdFixPrc)    
  if @row> 0     
  begin    select @errmsg= 'Enter Fix Price (Y/N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPRtn and OdGldAs<> '' and     
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdGldAs)           
  if @row> 0     
  begin    
   select @errmsg= 'Gold As Value Not Defined'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdLabAs)    
  if @row> 0     
  begin    
   select @errmsg= 'Labour As Value Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdIWtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. Equal To Order Wt. (Y/N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdIWtFrOrd<> '' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'IWTFRORD' and PMCd= OdIWtFrOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. From Order Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdIWtFrOrd<> '' and OdIWtEqOrd= 'Y')     
  if @row> 0     
  begin    
   select @errmsg= 'Mismatch in values InvWtEqOrd and InvWtFrOrd'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'INRTORD' and PMCd= OdIGldRtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invalid value for Gold Rate Equal To Order Rate'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and     
     not exists (select 'x' from Param where PTyp='YN' and PMCd= OdCls)    
  if @row> 0     
  begin    
   select @errmsg= 'Order Closed (Y/N)'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'WH' and PMCd= OdWh)    
  if @row> 0      
  begin    
   select @errmsg= 'W/H Set Option Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdWDiaAvlblDt in('01/01/1980', '01/01/1900') or     
   OdHDiaAvlblDt in('01/01/1980', '01/01/1900') or OdFndAvlblDt in('01/01/1980', '01/01/1900'))    
  if @row> 0     
  begin    
   select @errmsg= 'Wax/ Hand Set Diamond Availability/Fnd Availability Date Cannot Be Blank'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPRtn  and OdTc <> 'JRI' and OdTc <> 'JRO' and    
    not exists (select 'x' from Param where PTyp= 'CPYRT' and PMCd= OdGmChk)    
  if @row> 0     
  begin    
   select @errmsg= 'Gram Check Option Not Defined'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
    not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChkTol)    
  if @row> 0     
  begin    
   select @errmsg= 'Should Tolerence Be Checked? (Specify Y/ N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and     
    ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and     
    ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdDmCd= del.OdDmCd)     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design'    
   goto ErrHndlr    
  end  
  --5717
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
  --5717
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
   
  --5717
  select @Row= Count('x') from inserted join InvHd on InCoCd= OdCoCd and InTc= OdTc and   
   InYy= OdYy and InChr= OdChr and InNo= OdNo where InCls= 'Y' and OdTc='FOT'  
  If @Row > 0  
  begin   
	Select @ErrMsg= 'Cannot Add/Change As Invoice Is Closed'  
	GoTo ErrHndlr  
  end  
  
    
  /******  OdVaCtg  =  PMCd ('VACTG')  ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where 
   not exists (select 'x' from Param where PTyp= 'VACTG' and PMCd= OdVaCtg)    
  if @row> 0     
  begin    
   select @errmsg= 'Value Addition Category Not Defined'    goto ErrHndlr    
  end    
    
  /******  '  =  PMCd ('KT')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and    
   not exists (select 'x' from Param where PTyp= 'KT' and PMCd= OdKt)    
  if @row> 0     
  begin    
   select @errmsg= 'Karatage Not Defined'    
   goto ErrHndlr    
  end    
    
  Select @row= count('x') from inserted Ins where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and    
   (Ins.OdDmCol= '' or     
   ( Ins.OdDmCol <> IsNull((Select Del.OdDmCol From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
      Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy And     
      Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr ), '') And     
    not exists (select 'x' from Param where PTyp= 'DMCOL' and PMCd= OdDmCol And PValidYn = 'Y')     
   ))     
  if @row> 0     
  begin     
   select @errmsg= 'Design Colour Not Defined / Invalid.'     
   goto ErrHndlr     
  End     

  Select @row= count('x') from inserted where @wIsJLTc = 'N' and 
   not exists (select 'x' from vParam where vPCoCd= OdCoCd and vPTyp= 'LINE' and vPMCd= OdLine and vPMCd<> 'ZZZ')    
  if @row> 0    begin    
   select @errmsg= 'Line Not Defined/ Line Cannot Be ''ZZZ'' '    
   goto ErrHndlr    
  end    
    
  --5717 password needed to change prdseq when bag exist
  select @row= count('x') from  deleted del Join OrdMst on OmIdNo=del.OdOmIdNo and OmCoCd= del.OdCoCd and OmTc= del.OdTc  
       and OmYy= del.OdYy and OmChr= del.OdChr and OmNo= del.OdNo where 
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdPrdSeq= del.OdPrdSeq) 
	     and (exists(Select 'x' from Bag where BOdIdNo=del.OdIdNo))    
	     AND not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change PrdSeq When Bag Exists'    
   goto ErrHndlr    
  end   
    
  select @row= count('x') from  deleted del where 
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and     
     ins.OdDmSz= del.OdDmSz and ins.OdSfx= del.OdSfx and ins.OdOrdEnt= del.OdOrdEnt and ins.OdPrdEnt= del.OdPrdEnt) 
	     and (exists(Select 'x' from Bag where BOdIdNo=del.OdIdNo))    
	     

  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design Size/ Suffix/ OrdEnt/ PrdEnt When Bag Exists'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted ins     
 join OrdMst insOm On insOm.OmIdNo=ins.OdOmIdNo       
  where 
     (Select PValue from Param where PTyp='TC' and PMCd= ins.OdTc)= 'PL' and     
     (Select count('x') from OrdDsg Od join OrdMst Om on  Om.OmIdNo=Od.OdOmIdNo where 
		Om.OmCoCd= insOm.OmCoCd and Om.OmTc= (Case When insOm.OmCmCtg = 'S' Then 'SPL' Else 'PL' End)     
    and Om.OmCmCtg= insOm.OmCmCtg and Om.OmCmCd= insOm.OmCmCd and Od.OdDmCd= ins.OdDmCd     
    and Od.OdSfx= ins.OdSfx and Od.OdDmSz= ins.OdDmSz and Om.OmCmCurCd= insOm.OmCmCurCd) > 1     
  if @row> 0     
   begin    
     select @errmsg= 'Price List Already Defined for this Cust, Curr, Dsg, Sz, Sfx combination'    
     goto ErrHndlr    
   End    

  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdDelDt= '01/01/1900' or (OdDelDt<> '01/01/1980' and     
   OdDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc     
     and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo))     
  if @row> 0     
  begin    
   select @errmsg= 'Production Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr     
  end     

  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdExpDelDt= '01/01/1900' or (OdExpDelDt<> '01/01/1980' and     
   OdExpDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd     
    and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo))     
    
  if @row> 0     
  begin    
   select @errmsg= 'Export Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr   end    
    
  Select @row= count('x') from inserted where oDtC <> 'JCN' and @wIsJLTc = 'N' and OdOrdQty < 0 or OdPrdQty< 0    
  if @row> 0     
  begin    select @errmsg= 'Order/ Prd Qty Should Be >= 0 '    
   goto ErrHndlr    
  end    
    
 select @row= count('x') from inserted where odtc <> 'JCN' and @wIsJLTc = 'N' and 
   Round(OdOrdQty, 1) < Round(OdExpQty, 1)     
  
  if @row> 0     
  begin    select @errmsg= 'Order Qty Should Be >= Export Qty '    
   goto ErrHndlr    
  end    
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
  --5717
  --select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
  -- not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and    
  --   ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdPrtKey=del.OdPrtKey and ins.OdOrdEnt= del.OdOrdEnt and ins.OdPrdEnt= del.OdPrdEnt    
  --    and ins.OdPrdSeq= del.OdPrdSeq) and     
  -- (exists (Select 'x' from Bag where BOdIdNo=OdIdNo and BCoCd= del.OdCoCd and BOdTc= del.OdTc     
  --   and BOdYy= del.OdYy and BOdChr= del.OdChr and BOdNo= del.OdNo     
  --   and BOdSr= del.OdSr and BPrtKey=del.OdPrtKey) Or OdPrtFgQty<>0 or OdPrtExpQty<>0)    
  --if @row> 0     
  --begin    
  -- select @errmsg= 'Cannot Change OrdEnt/ PrdEnt / PrdSeq When Bag Exists'    
  -- goto ErrHndlr    
  --end  
  --5717
   select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdBagPcs<= 0 and OdTc <> 'JRI' and OdTc <> 'JRO'        
  if @row> 0     
  begin    
   select @errmsg= 'Default Bag Opening Pieces Cannot Be <= Zero'    
   goto ErrHndlr    
  end 

  select @row= count('x') from inserted where OdOmCtCd<> '' and     
   not exists (select 'x' from CustMst where CmCtg= 'T' and CmCd= OdOmCtCd)    
  if @row> 0     
  begin     
   select @errmsg= 'Contractor Code Not Defined'    
   goto ErrHndlr    
  end    
    
  
  select @row= count('x') from inserted ins     
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo 
  
  where OdGrMet<>''     
   and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'MET' and PSCd= OdGrMet)     
  if @row> 0     
  begin     
   select @errmsg= 'Metal Grade not Defined'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted ins    
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo  
  where OdGrDia<>'' and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'DIA' and     
   PSCd= OdGrDia)     
  if @row> 0     
  begin     
   select @errmsg= 'Diamond Grade not Defined'    
   goto ErrHndlr    
  end    
    
  if @wTcTyp = 'JT'   
	  Begin   
	  -- rewrite 

	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmIdNo=Ins.OdOmIdNo 
	Where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
			Where Od.OdCoCd=(Case when ins.OdTc in ('JIN','LIN')  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
			
	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmIdNo=Ins.OdOmIdNo 
	Where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       ((TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
			
			Where Od.OdCoCd=Ins.OdCoCd  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo Or Od.OdDmStkNo=0)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))  
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
   
   Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo 
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
          
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
            Where Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1 =Ins.OdJLLn1 and
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.OdJLLn3    
            and (Od.OdDmStkNo=Ins.OdDmStkNo/*Or (Od.OdDmStkNo=0 Or Ins.OdDmStkNo=0)*/)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))
			
	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo 
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
	   (
       (TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
            Where Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and (Od.OdJLLotNo=Ins.OdJLLotNo or Od.OdJLLotNo = '') and Od.OdJlLn1 = Ins.OdJlLn1 and 
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))  
       )  
        )  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
   Select @Row=Count('x') from Inserted Ins Where 
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
           ), '')+',%')  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  
	Select @Row=Count('x') from Inserted Ins where
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.ODJLRmCd =Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1=Ins.OdJLLn1 
				and Od.ODJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
           ), '')+',%')  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  

	Select @row= count('x') from inserted where 
    exists(Select 'x' from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT') and Round(OdPrdQty, 1) > 0    
	   	   	   	   if @row> 0     
    begin    select @errmsg= 'Prd Qty should be 0 for Trading voucher'    
     goto ErrHndlr    
    end    

     
  end  
  

End    
--5717
if @mode = 'C'    
	  Begin      
	  select @Row= Count('x') from deleted del where     
		not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
		and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
		and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr)    
	  If @Row > 0    
	  begin     
	   Select @ErrMsg= 'Cannot Change Company Code/Tc/Yy/Chr/No/Sr'    
	   GoTo ErrHndlr    
	  end    
	 --5717
	  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and   
	  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChgPtrOnStwYN and PSCd= '')  
	 if @row> 0     
	 begin  
	  select @errmsg= 'Change Pointer On STW (Y/N)'  
	  goto ErrHndlr  
	 end   
	 --5717
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
	 --5717
	 select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and  
	  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdILabWtFrOrd and PSCd= '')  
	 if @row> 0     
	 begin  
	  select @errmsg= 'Labor Weight From Order (Y/N)'  
	  goto ErrHndlr  
	 end  
	 --5717
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
end    

 if @mode = 'D'   
 Begin    
  Select @row= count('x') from deleted where 
   (exists(Select 'x' from Bag where     
                BOdIdNo=OdIdNo ))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Export Has Taken Place / Design Under Production'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and      	 
    exists(Select 'x' from OrdLn where OdlCoCd =OdCoCd and OdlTc =OdTc and OdlYy =OdYy and OdlChr =OdChr 
    and OdlNo =OdNo and OdlSr =OdSr and OdlOdIdNo =OdIdNo )   
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Transaction Exist in OrdLn'    
   goto ErrHndlr    
  end   
  --5717
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and    
   (OdPrtExpQty<>0 Or exists(Select 'x' from Fgd where FdCoCd= OdCoCd and FdExpOdTc= OdTc and     
         FdExpOdYy = OdYy and FdExpOdChr = OdChr and FdExpOdNo = OdNo and     
         FdExpOdSr = OdSr and FdPrtKey=OdPrtKey))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Record Exists As Export Order In Fgd Table'    
   goto ErrHndlr    
  end 
    
  if @wTcTyp = 'JT'   
  Begin   
  
  /* Cannot Delete JRI if exist in Txnd */
  select @row= count('x') from deleted del where del.OdTc='JRI' and del.OdDmStkYy <> ''  and del.OdDmStkChr <> '' and del.OdDmStkNo <> ''  
    and exists(Select 'x' from Txnd tx where tx.TdCoCd =del.OdCoCd and tx.TdBYy = del.OdDmStkYy and tx.TdBChr = del.OdDmStkChr and 
		tx.TdBNo = del.OdDmStkNo and tx.ModDt >= del.ModDt and tx.ModTime >= del.ModTime)       
  if @row > 0    
  begin    
   select @errmsg= 'Cannot Delete, Bag Transaction Exists for this Repair Bag'    
   goto ErrHndlr    
  end      
  /* why is OdBaseIdKey required to be checked index is there or not */
  Select @row=Count('x') from deleted Del Where 
    exists(Select 'x' from OrdDsg Od Where   
       (Od.OdBaseIdKey=Del.OdIdNo))  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Base Voucher'  
    goto ErrHndlr  
   End  
    
   Select @row=Count('x') from deleted Del Where 
    exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and 
       Od.OdRefIdKey=Del.OdIdNo)  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Reference Voucher'  
    goto ErrHndlr  
   End  
  End   
 End    
   
 /* **** Manali Trading Module **** */  
 if @mode ='C'  
 Begin  
 if @wTcTyp = 'JT'   
  Begin  
  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where @wIsJLTc = 'N' and
   exists(Select 'x' from OrdDsg Od Where 
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
    and ins.OdDmCd=del.OdDmCd and ins.OdSfx=del.OdSfx and ins.OdDmSz=del.OdDmSz   
	and ins.OdDmStkYy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr 
    and ins.OdDmStkNo=del.OdDmStkNo
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  
      
  Select @row=Count('x') from deleted Del Where 
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and 
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'N' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and 
   ((ins.OdDmStkyy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr and ins.oddmstkno = del.oddmstkno and ins.OdDmStkNo=del.OdDmStkNo) or ins.OdDmStkNo=0 or del.OdDmStkNo=0))    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  

 End  

  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where @wIsJLTc = 'Y' and
   exists(Select 'x' from OrdDsg Od Where  (Del.OdTc='LIN' Or  Od.OdCoCd=Del.OdCoCd) and   
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
    and ins.OdJLRmCd=del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJLLn2 = del.odJLLn2 and ins.odJLLn3 = del.odJLLn3 
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  

    
  Select @row=Count('x') from deleted Del Where   
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and   
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'Y' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
     and ins.OdJLRmCd =del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJlLn2 = del.OdJlLn2 and ins.OdJLLn3 = del.ODJLLn3
     )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  
 End  

 --??End  
  
 if @mode in ('C','D')  
 begin  
 If @wTcTyp = 'FG' Or @wTcTyp = 'JT'         
 BEgin  
  update InvHd set   
  @wInFob= convert (decimal (16, 2), InFOB- IsNull((Select sum((del.OdSalPrc*del.OdOrdQty)+del.OdJLVchVal) from deleted del where del.OdCoCd= InCoCd and del.OdTc= InTc and del.OdYy= InYy   
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
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo   
  where InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD In ('FG', 'JT') and PValue= OdTc)  
         or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  
  if @wIsJLTc = 'N'
  begin
  
  update DmYyMm set  
  DDmCrQty= DDmCrQty- (select isnull(sum(Del.OdOrdQty), 0) from Deleted Del  
   Join OrdMst on   
   OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm),  
  DDmDrQty= DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkyy = DDmStkYy and del.OdDmStkChr = 
			DDmStkChr and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd    
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkCoCd =DDmStkCoCd and Deleted.OdDmStkNo= DDmStkNo and Deleted.OdDmStkYy= DDmStkYy and Deleted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
   else
   begin
  update DmYyMm set  
    DDmCrWt= DDmCrWt- (select isnull(sum(Del.OdJLWt), 0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm),  
  DDmDrWt= DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd= DJLRmCd 
   and Deleted.OdJLLotNo= DJLLotNo and Deleted.OdJLLn1= DJLLn1 and Deleted.OdJLLn2= DJLLn2 and Deleted.OdJLLn3 = DJLLn3 and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
 End  
  end
    
	if @mode in ('A','C')
	begin
		Select @row= count('x') from inserted where Inserted.OdTc='JRI' and 
			exists (Select 'a' from Bag where BCoCd=OdDmStkCoCd and BYy=OdDmStkYy and BChr = OdDmStkChr and BNo=OdDmStkNo) and
				not exists 
				(Select 'a' from Bag where 
							BCoCd = OdDmStkCoCd and BOdDmCd=OdDmCd and BOdSfx=OdSfx and BOdDmTcTyp = 'DM' and
							BOdDmSz=OdDmSz and BYy = OdDmStkYy and BChr =OdDmStkChr and BNo = OdDmStkNo and BCls='Y' and BLoc='PFG' and BRepYN='N')
 
		if @row>0 
		begin
			select @errMsg = 'Given DmStkNo not a valid Bag No For this Dsg Code-Sz-Suffix or Bag not closed or Bag not in FG or Bag already in Repair'
			goto ErrHndlr
		end
	end 
 
 If @mode in ('A','C')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'   
 Begin  

  /*** Update InvHd ***/  
  update InvHd set  
  @wInFob = convert (decimal (16, 2), InFOB+ IsNull((Select sum((ins.OdSalPrc*ins.OdOrdQty)+ins.OdJLVchVal) from inserted ins where   
   ins.OdCoCd= InCoCd and ins.OdTc= InTc and ins.OdYy= InYy   
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
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo   
  where InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD ='FG' and PValue= OdTc)  
        or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  

  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr, DDmStkNo, DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt,DdmStkCoCd)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy,OdDmStkChr,OdDmStkNo,OmFrDmLoc,convert(int,convert(varchar(6),OdOmDt,112)),0,0,@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0,OdDmStkCoCd   
   from inserted join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DCoCd= OdCoCd and DDmCd=OdDmCd   
        and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy= DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmFrDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and OdDmStkCoCd =DdmStkCoCd and 
        DYyMm= convert(int,convert(varchar(6),OdOmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
         
  /* Current Partition */   
  if @wIsJLTc ='N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd =DDmStkCoCd and Ins.OdDmStkNo = DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd =DDmStkCoCd and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdDmCd= DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkCoCd =DDmStkCoCd  
   and Inserted.OdDmStkYy = DDmStkYy and Inserted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdCoCd= DCoCd 
   and OmFrDmLoc=DLoc  and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd= DJLRmCd 
   and Inserted.OdJLLotNo= DJLLotNo and Inserted.OdJLLn1= DJLLn1 and Inserted.OdJLLn2= DJLLn2 and Inserted.OdJLLn3 = DJLLn3  and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 end  
  end
  
 if @mode in ('C','D')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'     
 Begin  
 if @wIsJLTc = 'N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst  
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst   
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd   
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkCoCd =DDmStkCoCd and Deleted.OdDmStkNo = DDmStkNo and Deleted.OdDmStkYy = DDmStkYy and 
   Deleted.OdDmStkChr = OdDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
  else
  update DmYyMm set  
  DDmCrWt=DDmCrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst  
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst   
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd = DJLRmCd and Deleted.OdJLLotNo = DJLLotNo and Deleted.OdJLLn1 = DJLLn1 and
   Deleted.OdJlLn2 = DJLLn2 and Deleted.OdJllN3 = DJLLn3  and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
 If @wTcTyp = 'JT'     
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty- IsNull((Select sum(del.OdOrdQty) from deleted del where   
     (del.OdTc='JIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt- IsNull((Select sum(del.OdJLWt) from deleted del where   
    (del.OdTc='LIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from deleted where   
    OrdDsg.OdCoCd= deleted.OdCoCd and OrdDsg.OdIdNo=deleted.OdBaseIdkey  
   and Exists (Select 'x' from TcMst Where TmCoCd=deleted.OdCoCd and TmTc=deleted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls = (case when @wIsJLTc ='N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
				else (case when OrdDsg.OdJLWt= OrdDsg.OdJLExpWt and OrdDsg.OdJLWt<> 0 then 'Y' else 'N' end)   end)
   from deleted del where OrdDsg.OdCoCd= Del.OdCoCd and OrdDsg.OdIdNo= Del.OdBaseIdKey  
      
   and Exists (Select 'x' from TcMst Where TmCoCd=del.OdCoCd and TmTc=del.OdTC and TmBaseTCs<>'')  
 
  If @mode in ('D','C')
  Begin  
   Update Bag set
		BRepYN = 'N',   
		BRepQty = 0,
		BRepINDt                ='01/01/1980',
		BRepLoc  = '',
		BRepGrWt	= 0,
		BRepOdTc = '',
		BRepOdYy = '',
		BRepOdChr = '',
		BRepOdNo = 0,
		BRepOdSr = 0
   from Deleted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRI'

   Update Bag set
   		BRepYN = 'Y'
   from Deleted
		where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRO'

  End    

 End   
  
 End   
  
 if @mode in ('A','C')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'    
 Begin  
  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr,DDmStkNo,DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt,DdmStkCoCd)
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy, OdDmStkChr, OdDmStkNo,OmToDmLoc,convert(int,convert(varchar(6),OdOmDt,112)),0,0,
					@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0,OdDmStkCoCd  
     
   from inserted join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DCoCd= OdCoCd and DDmCd=OdDmCd   
       and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy = DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmToDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and OdDmStkCoCd =DdmStkCoCd and 
       DYyMm= convert(int,convert(varchar(6),OdOmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  
  if @wIsJLTc = 'N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd= DDmStkCoCd and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd= DDmStkCoCd and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdDmCd=DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkCoCd= DDmStkCoCd and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkChr = DDmStkChr and Inserted.OdDmStkYy = DDmStkYy 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd = DJLRmCd and Inserted.ODJLLotNo = DJLLotNo and Inserted.ODJLLn1 = DJLLn1 and Inserted.ODJLLn2 = DJLLn2 and 
   Inserted.ODJLLn3 = DJLLn3 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 If @wTcTyp = 'JT'   
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty + IsNull((Select sum(Ins.OdOrdQty) from Inserted Ins where   
   (ins.OdTc='JIN' Or Ins.OdCoCd= OrdDsg.OdCoCd)  and Ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt+ IsNull((Select sum(ins.OdJLWt) from Inserted Ins where   
   (ins.OdTc='LIN' Or ins.OdCoCd= OrdDsg.OdCoCd) and ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from Inserted where 
    OrdDsg.OdCoCd= Inserted.OdCoCd and OrdDsg.OdIdNo=Inserted.OdBaseIdkey  
    and Exists (Select 'x' from TcMst Where TmCoCd=Inserted.OdCoCd and TmTc=Inserted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls =(case when @wIsJLTc = 'N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
			else (case when OrdDsg.OdJLWt = OrdDsg.ODJLExpWt and OrdDsg.ODJLWt <> 0 then 'Y' else 'N' end) end)
   from Inserted Ins where (ins.OdTc = 'LIN' or ins.OdTc='JIN' Or  OrdDsg.OdCoCd= Ins.OdCoCd)   
   and OrdDsg.OdIdNo= Ins.OdBaseIdKey and Exists (Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTC and TmBaseTCs<>'')  
  if @mode in ('A','C')  
  Begin   
   Update Bag set   
		BRepYN = 'Y',
		BRepQty = OdOrdEnt,
		BRepINDt = getdate(),
		/* (Select OmDt From OrdMst Where OmCoCd= OdCoCd And OmTc = OdTc   
		And OmYy = OdYy And OmChr= OdChr And OmNo= OdNo And OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo)*/
		BRepLoc  = (Select HDefRepBagLoc from Head where HCoCd = Inserted.OdCoCd),
		BRepGrWt = OdRepGrWt,
		BRepOdTc = OdTc,
		BRepOdYy = OdYy,
		BRepOdChr = OdChr,
		BRepOdNo = OdNo,
		BRepOdSr = OdSr
   from Inserted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRI'

   Update Bag set
		BRepYN= 'N'
   from Inserted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRO'
  End  
 End  
  
 End    
  
  
if @mode in ('C','A') and @wTcTyp = 'JT'   
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') from Inserted Ins   
   where exists(Select 'x' from OrdDsg B Where  
      B.OdIdNo=Ins.OdBaseIdkey and B.OdOrdQty < B.OdExpQty)  
    and (exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>''))  
  If @Row> 0  
  begin   
   select @errmsg= 'Order Qty Should Be >= Matched Qty for Base Voucher'    
   goto ErrHndlr    
  end
 end
else
begin
 Select @Row=Count('x') from Inserted Ins   
   where 
    exists(Select 'x' from OrdDsg B Where  
      B.OdIdNo=Ins.OdBaseIdkey and cast(B.OdJLWt as decimal(10,3))< cast(B.OdJLExpWt as decimal(10,3)))  
    and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>'')  
 If @Row> 0  
  begin   
   select @errmsg= 'Order Wt Should Be >= Matched Wt for Base Voucher'    
   goto ErrHndlr    
  end
 end   
end    
  
If @mode in ('C','A') and  (@wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then 
	  (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) 
	  else (case when OmFrDmDc = 'C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Ins.OdCoCd= DCoCd and Ins.OdDmCd= DDmCd and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and   
    Ins.OdDmStkCoCd =DdmStkCoCd and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkNo = DDmStkNo and 
     (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Desing Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
end
else
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrWt)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.OdJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3=DJLLn3 and
	(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Size'  
     GoTo ErrHndlr  
     end  

end
end  
If @mode in ('D') and  (@wTcTyp = 'FG' or @wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 -- rewrite checking JCN why and remove FG and see what to simplify 
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd = Del.OdCoCd and LocTyp =(Case PValue When 'JT' Then 'S' End)  and   
            LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then (case when OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and   
    Del.OdDmStkCoCd =DdmStkCoCd and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo and 
    (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Design Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
	end
else
begin
 -- rewrite checking JCN why and remove FG and see what to simplify 
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Del.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End)  and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrwT)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Del.OdCoCd= DCoCd and Del.ODJLRmCd = DJLRmCd and Del.ODJLLotNo = DJLLotNo and Del.ODJLLn1 = DJLLn1 and Del.ODJLLn2 = DJLLn2 and Del.ODJLLn3 = DJLLn3 and
(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Sz'  
     GoTo ErrHndlr  
     end  
end

end  
return    
ErrHndlr:    
 begin   select @errmsg='|'+@errmsg    
  raiserror(@errmsg,16,1)   rollback transaction     
 end    
end   



GOGO
ALTER TRIGGER [dbo].[DsgRm_INUPDL] ON [dbo].[DsgRm] 
       FOR INSERT,UPDATE,DELETE  
	   NOT FOR REPLICATION
AS  
BEGIN  
	/* Trigger Version 3.02.0 - Speed Optimization */
	DECLARE @ROW INT  
	DECLARE @mode char(1)  
	DECLARE @ReplAgt BIT  
	Declare @wHCd Varchar(8)       
	DECLARE @errmsg varchar(255) 
	DECLARE @SelfCoCd varChar(3) 
	DECLARE @wCurrPrtn VarChar(1)
	
	IF @@ROWCOUNT=0   RETURN  
	
	/* Sachin 3.02 - Constant For Current Partition */
	select @wCurrPrtn = 'C'		
	Select @SelfCoCd ='ZZZ'
	Select @wHCd = 'ZSELF'

	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'DsgRm' )  
	RETURN  
	
	select @mode=case (select count('x') from inserted) when 0 then 'D'  
	else  
	case (select count('x') from deleted) when 0 then 'A' else 'C' end  
	end   
	
	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))
	
	/******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */
	if @mode='A'  And @ReplAgt <> 1
	begin
		Select @row=count('x') From Inserted Where DrPrtKey <> @wCurrPrtn 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Insert Partition Key Other Than ''C'''
			goto ErrHndlr
		end
	end

	if @mode='C'  And @ReplAgt <> 1
	begin
		Select @row=count('x') From Inserted Where DrPrtKey<>@wCurrPrtn
		if @row> 0 
		begin
			select @errmsg= 'Cannot Edit Records of Previous Partition'
			goto ErrHndlr
		end
	end 

	If @Mode='D'  And @ReplAgt <> 1
	Begin 
			select @row= count('x') from deleted del where del.DrPrtKey<>@wCurrPrtn
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete Records From Previous Partition'
				goto ErrHndlr
			end
	End
	/******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */	

	If  (@ReplAgt<>1)
	begin
	 	/******Cannot Insert/Update/Delete if Crm database /Emr database with sdMst='N' ****/
		if @mode in ('A','C','D')
		Begin 
		If exists( select 'x' From SysDb where SysId='EMR' and SdMst='N' )  or  
		exists ( select 'x' From SysDb where SysId='CRM' and SysDb ='WEB' ) 
			Begin 
				select @errmsg= 'Cannot Insert /Update /Delete ' 
				goto ErrHndlr 
			end 
		End 
	End		
	/******** Key Cannot Be Changed ********/
 	if @mode = 'C'
 	Begin
		/* Prt Key Not required as only inserted and deleted record is compared */
  		select @Row=Count('x') from deleted del where			
    		not exists (Select 'x' from inserted ins where ins.DrIdNo=del.DrIdNo and ins.DrTcTyp= del.DrTcTyp 
				and ins.DrCd= del.DrCd and ins.DrSz= del.DrSz and ins.DrSr= del.DrSr)
  		If @Row > 0
	  	begin
			Select @ErrMsg='Cannot Change DsgRm Key'
			GoTo ErrHndlr
		end
	End
	/****** DrTcTyp, DrCd, DrSz combination should exist in the DsgMst ******/
	/* ****** Sachin 3.02 - Currrent Partition	*/
	select @row= count('x') from inserted Ins where 
			not exists (select 'x' From Deleted Del Where Del.DrIdNo=Ins.DrIdNo and Del.DrTcTyp= Ins.DrTcTyp and Del.DrCd= Ins.DrCd 
					and Del.DrSz= Ins.DrSz and Del.DrPrtKey<>Ins.DrPrtKey) and  
			not exists (select 'x' from DsgMst where 
				DmIdNo=Ins.DrDmIdNo and DmTcTyp= Ins.DrTcTyp and DmCd= Ins.DrCd and DmSz= Ins.DrSz and DmPrtKey=Ins.DrPrtKey)
	if @row> 0
	Begin
		select @errmsg= 'This Type-Design-Size Combination Not Defined In Design Master'
		goto ErrHndlr
	end
	/******  DrWt  <>  0  ******/
	select @row= count('x') from inserted where Round(DrWt, 3) = 0
	if @row> 0
	Begin
		select @errmsg= 'Raw Material Weight Cannot Be Zero'
		goto ErrHndlr
	end
	/******  DrRmCtg  =  PMCd ('RMCTG')  ******/
	select @row= count('x') from inserted where
		not exists (select 'x' from Param where
		PTyp= 'RMCTG' and PMCd= DrRmCtg)
	if @row> 0
	Begin
		select @errmsg='Raw Material Category Not Defined'
		goto ErrHndlr
	end
	/******  DrRmSCtg  =  PSCd ('RMCTG', DrRmCtg)  ******/
	select @row=count('x') from inserted where
		not exists (select 'x' from Param where
				PTyp='RMSCTG' and PMCd=DrRmCtg
				and PSCd=DrRmSCtg)
	if @row>0
	Begin
		select @errmsg='Raw Material Sub Category Not Defined For This Category'
		goto ErrHndlr
	end
/*  ######################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ######################  */
	/******  DrRmCd  =  valid RmCd from RmMst (RmCtg= DrRmCtg, RmSCtg= DrRmCtg, RmCd= DrRmCd, RmZ= 'N')  ******/
	/* ****** Sachin 3.02 - Current Partition - CHECK REQUIRED	*/
	SELECT @row= count('x') FROM Inserted Ins WHERE Ins.DrPrtKey=@wCurrPrtn and
		Not Exists (Select 'x' From Deleted Del Where Ins.DrPrtKey<>Del.DrPrtKey and Ins.DrDmIdNo=Del.DrDmIdNo 
			and Ins.DrTcTyp = Del.DrTcTyp And Ins.DrCd = Del.DrCd AND Ins.DrSz = Del.DrSz) and 
		(Ins.DrRmCd <> ISNULL((SELECT Del.DrRmCd FROM Deleted Del WHERE Ins.DrPrtKey=Del.DrPrtKey and Ins.DrIdNo=Del.DrIdNo 
			and Ins.DrTcTyp = Del.DrTcTyp And Ins.DrCd = Del.DrCd AND Ins.DrSz = Del.DrSz AND Ins.DrSr = Del.DrSr), '')
		AND NOT EXISTS (SELECT 'x' FROM RmMst WHERE	RmCtg= DrRmCtg and RmSCtg= DrRmSCtg
			And RmCd= DrRmCd and RmZ= 'N' And RmValidYN = 'Y' And RmPrtKey=DrPrtKey))
		OR DrRmCd =''
	if @row> 0
	Begin
		select @errmsg= 'Raw Material Code Not Defined/InValid'
		goto ErrHndlr
	end
/*	######################  Manoj #### Ver: 2.0.9 #### Date: 27/12/2004  ######################  */
	/* ****** Sachin 3.02 - Current Partition - CHECK REQUIRED	*/
	select @row= count('x') from inserted Ins where Ins.DrPrtKey=@wCurrPrtn and
		Not Exists (Select 'x' From Deleted Del Where Ins.DrDmIdNo=Del.DrDmIdNo and Ins.DrTcTyp=Del.DrTctyp and Ins.DrCd=Del.DrCd 
					and Ins.DrSz=Del.DrSz and Ins.DrPrtKEy<>Del.DrPrtKey) and
		not exists (select 'x' from RmMst where
				RmCtg= Ins.DrRmCtg and RmSCtg= INs.DrRmSCtg and RmCd= Ins.DrRmCd and RmPrtKey=Ins.DrPrtKey and RmZ= 'N')
	if @row> 0
	Begin
		select @errmsg= 'Raw Material Code Not Defined'
		goto ErrHndlr
	end
/*	##########################################  Manoj ##########################################  */

/* *** Bef 2.14Next *** */
	/* *** Jay 2.13(CT) ***  		(new CT) */
	/* Can Not Insert > 1, RmCtg  in (G,P,S) For DmTcTyp = 'PM'  when Module is 'Y'*/
	/*
	select @row= count('x') from inserted Ins where  DrTcTyp = 'PM' and 
		(select Count('x') from DsgRm Dr where ins.DrTcTyp = Dr.DrTcTyp  and ins.DrCd= Dr.DrCd and
		ins.DrSz= Dr.DrSz and DrRmCtg in ('G','P','S')) > 1
		and exists (select 'x' from Head Where hCompBagsYn ='Y' and HCoCd= 'ZZZ')
	if @row> 0
	Begin
		select @errmsg= 'Can Not Insert More Then One Metel For The Component'
		goto ErrHndlr
	end
	*/
	/* *** Jay 2.13(CT) ***  		(new CT) */
/* *** Bef 2.14Next *** */

	/******  DrSetSCd  =  PSCd ('LABSCD', 'SET')  ******/
	SELECT @row= count('x') FROM inserted Ins WHERE DrRmCtg in ('D', 'C') AND
		((Ins.DrSetSCd <> ISNULL((SELECT Del.DrSetSCd FROM Deleted Del WHERE 
			Ins.DrIdNo=Del.DrIdNo And Ins.DrTcTyp = Del.DrTcTyp And Ins.DrCd = Del.DrCd AND Ins.DrSz=Del.DrSz AND Ins.DrSr = Del.DrSr), '')
		AND NOT EXISTS (SELECT 'x' FROM Param WHERE PTyp= 'LABSCD' and PMCd= 'SET'
			AND PSCd= DrSetSCd And PValidYN = 'Y'))
		OR DrSetSCd = '')
	if @row> 0
	Begin
		select @errmsg= 'Setting Code Not Defined For This Rm Category'
		goto ErrHndlr
	end

    /****** Sachin 3.3.0 (30-06-08) - [28. Valid OrLn1/OrLn2 Check]   ******/      
    Select @row=Count('x') From Inserted Where @mode In ('A','C') and (update(DrLn1) Or Update(DrLn2) Or Update(DrRmCd)) And DrPrtKey=@wCurrPrtn And DrRmCtg In ('D','C') And
			Exists(Select 'x' From Head Where HCoCd=@SelfCoCd and HCd= @wHCd And hChkVldRmLnYN='Y') And
			Not Exists(Select 'x' From VldRmLn Where VrsRmCtg=DrRmCtg 
								And VrsRmId=(Select RmId From RmMst 
									Where RmCtg=DrRmCtg And RmCd=DrRmCd) 
												And VrsLn1=DrLn1 And VrsLn2=DrLn2)		 
    if @row>0       
    Begin      
	   select @errmsg='Ln Not Defined/Ln1 And Ln2 Combination Not Defined In RM Ln Master'      
	   goto ErrHndlr      
	end      

/* (Bef 2.06)
	select @row= count('x') from inserted where DrRmCtg in ('D', 'C') and
		not exists (select 'x' from Param where PTyp= 'LABSCD' and PMCd= 'SET' and PSCd= DrSetSCd)
	if @row> 0
	Begin
		select @errmsg= 'Setting Code Not Defined For This Rm Category'
		goto ErrHndlr
	end
*/
/*  #########################################  Manoj  #########################################  */
	/******  If  DrRmCtg = 'C' or 'D' then  DrSetSCd  =  PSCd ('LABSCD', 'SET')  	else 	DrSetSCd  =  Blank ******/
	--5717 subshp removed
	select @row= count('x') from inserted where DrRmCtg not in ('D', 'C') and (DrSetSCd<> '' )
	if @row> 0
	Begin
		select @errmsg= 'Setting Code Should Not Be Defined For This Rm Category'
		goto ErrHndlr
	end
	/******  DrSubShp  =  Blank  or  PMCd ('SUBSHP')  ******/
	select @row= count('x') from inserted where DrSubShp<> '' and
		not exists (select 'x' from Param where
				PTyp= 'SUBSHP' and PMCd= DrSubShp)
	if @row> 0
	Begin
		select @errmsg= 'Sub Shape Not Defined'
		goto ErrHndlr
	end
	/******  DrAlyCd  =  Blank  or  valid RmCd from RmMst (RmCtg= 'A')  ******/
	/* ****** Sachin 3.02 - Current Partition	*/
	/* ****** Manali 3.03 - 18/06/08 - 'S' and 'L' added */
	select @row= count('x') from inserted where DrRmCtg in ('G', 'P', 'S', 'L') and
		DrAlyCd<> '' and not exists (select 'x' from RmMst where
						RmCtg= 'A' and RmCd= DrAlyCd and RmPrtKey=@wCurrPrtn)
	if @row> 0
	Begin
		select @errmsg= 'Alloy Code Not Defined'
		goto ErrHndlr
	end
	/******  If DrRmCtg<> 'G' or 'P' then 	DrAlyCd  =  Blank  ******/
	/* ****** Manali 3.03 - 18/06/08 - 'S' and 'L' added */
	select @row=count('x') from inserted where DrRmCtg not in ('G', 'P', 'S', 'L')
		and DrAlyCd<>''
	if @row>0
	Begin
		select @errmsg='Alloy Code Should Not Be Entered For This Raw Material Category'
		goto ErrHndlr
	end
	/******  DrMainMet  =  PMCd ('YN')  ******/
	select @row= count('x') from inserted where
		not exists (select 'x' from Param where
			PTyp= 'YN' and PMCd= DrMainMet)
	if @row> 0
	Begin
		select @errmsg= 'Raw Material Code Is The Main Metal (Y/N)'
		goto ErrHndlr
	end
	/****** Cannot have more than one record with DrMainMet = 'Y'  for a Design Type, Code, Size combination  ******/
	/* ****** Sachin 3.02 - Current Partition	*/
	select @row= count('x') from inserted ins where ins.DrPrtKey=@wCurrPrtn and
		(select count('x') from DsgRm Dr where ins.DrTcTyp = Dr.DrTcTyp  and ins.DrCd= Dr.DrCd and
					ins.DrSz= Dr.DrSz and ins.DrPrtKey=Dr.DrPrtKey and Dr.DrMainMet = 'Y') > 1
	if @row> 0
	Begin
		select @errmsg= 'Cannot Have More Than One Main Metal'
		goto ErrHndlr
	end
return
ErrHndlr:
	begin
		select @errmsg='|'+@errmsg
		raiserror(@errmsg,16,1)
		rollback transaction
	end
end

GOGO
ALTER TRIGGER [dbo].[OrdRm_INUPDL] ON [dbo].[OrdRm]       
       FOR INSERT,UPDATE,DELETE        
		NOT FOR REPLICATION
AS        
BEGIN        
 /* Trigger Version 3.8.0  */      
 DECLARE @ROW INT        
 DECLARE @mode char(1)        
 DECLARE @errmsg varchar(255)        
 Declare @wSysDb Varchar(3)      
 Declare @wHCd Varchar(8)       
 DECLARE @ReplAgt BIT        
 DECLARE @wCurrPrtn VarChar(1)      
       
       
 IF @@ROWCOUNT=0   RETURN        
       
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'OrdRM' )        
 RETURN        
       
 select @mode=case (select count('x') from inserted) when 0 then 'D'        
  else        
   case (select count('x') from deleted) when 0 then 'A' else 'C' end        
  end         
      
      
 Select @wHCd= 'ZSELF'       
 Select @wCurrPrtn = 'C'       
      
 /* **** Checking Whether a Normal Transaction or Replicated One **** */      
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))      
 Select @wSysDb=SysDb From SysDb where SysId='CRM'      
 Set @wSysDb=IsNull(@wSysDb,'')       
       
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */      
 if @mode='C'      
 begin      
  Select @row=count('x') From Inserted Where OrPrtKey<>@wCurrPrtn      
  if @row> 0       
  begin      
   select @errmsg= 'Cannot Edit Records of Previous Partition'      
   goto ErrHndlr      
  end      
 End      
      
 if @mode='A'      
 Begin      
  Select @row=count('x') From Inserted Where OrPrtKey <> @wCurrPrtn       
  if @row> 0       
  begin      
   select @errmsg= 'Cannot Insert Partition Key Other Than ''C'''      
   goto ErrHndlr      
  end      
 end       
      
 If @Mode='D'       
 Begin       
  select @row= count('x') from deleted del where del.OrPrtKey<>@wCurrPrtn      
  if @row> 0       
  begin      
   select @errmsg= 'Cannot Delete Records From Previous Partition'      
   goto ErrHndlr      
  end      
 End      
 /******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */      
      
 If  (@ReplAgt<>1)      
 Begin      
      
  /*db='WEB' then cannot  Insert /Update/Delete  where OmChr<> 'WEB'       
  Cannot edit/Delete where OmSubmitYN=’Y’ i.e Check deleted.OmSubmitYN=’Y’        
  Cannot edit OmAckYN        
  */      
  If @Mode in ('A','C' )      
  Begin      
   /******Cannot Insert/Update/Delete  if Crm database /EmrDemo database with TrdCocd not in sdCoCd ****/      
   /***if SysId=EmrDemo Record not found then normal behaviour i.e can insert/update/delete****/      
   /* Current Partition */      
   Select @ROW =Count('x')  From inserted where  OrPRtKey=@wCurrPrtn and      
   Exists (Select 'x' From SysDb where (SysId='EmrDemo' and ',' + SdCoCd + ',' Not  like '%,'+ OrCoCd +',%'))      
   If @Row > 0        
   Begin       
    select @errmsg= 'Cannot Insert /Update '       
    goto ErrHndlr       
   end       
      
   /*****If  db= '-' then Insert  Not allowed in OrdDsg  for OmChr='WEB' if OmAckYN<>'Y' ******/      
   /* Current Partition */      
   Select @Row=Count('x') From Inserted where OrPrtKey=@wCurrPrtn and OrChr='WEB' and @wSysDb='-' and       
   (Select OmAckYN From OrdMst where OmCoCd=OrCoCd and OmTc=OrTc and OmYy=OrYy and OmChr=OrChr and OmNo=OrNo and OmPrtKey=OrPrtKey)='N'      
        
   If @Row> 0      
     Begin        
      select @errmsg= 'Cannot Insert/Update Order Rm Of an UnAcknowledged Web Orders'        
      goto ErrHndlr        
     end        
       
   /*****If  db= 'WEB' then Insert/Update  Not allowed in OrdMst  for OdChr<>'WEB'/OdChr='WEB' and OmSubmitYN='Y' ******/      
   /* Current Partition */      
   Select @Row=Count('x') From Inserted where OrPrtKey=@wCurrPrtn and @wSysDb='WEB' and       
   (OrChr<>'WEB' Or (OrChr='WEB' and (Select OmSubmitYN From OrdMst where OmCoCd=OrCoCd and OmTc=OrTc and       
        OmYy=OrYy and OmChr=OrChr and OmNo=OrNo and OmPrtKey=OrPrtKey)='Y')      
    )       
   If @Row> 0      
     Begin        
      select @errmsg= 'Cannot Insert/Update Ord Rm Other than Chr=''WEB'' and Submitted=''N'' '        
      goto ErrHndlr        
     end        
  End      
       
  If @Mode ='D'       
  Begin      
   /* Current Partition */      
   select @ROW =Count('x')  From Deleted where  OrPrtKey=@wCurrPrtn and      
    Exists (select 'x' From SysDb where (SysId='EmrDemo' and ',' + SdCoCd + ',' Not  like '%,'+ OrCoCd +',%'))      
   if @Row > 0        
   Begin       
    select @errmsg= 'Cannot Delete '       
    goto ErrHndlr       
   end       
      
   /*****If  db= '-' then Insert  Not allowed in OrdDsg  for OmChr='WEB' if OmAckYN<>'Y' ******/      
   /* Current Partition */      
   Select @Row=Count('x') From Deleted where OrPrtKey=@wCurrPrtn and  OrChr='WEB' and       
   (Select OmAckYN From OrdMst where OmCoCd=OrCoCd and OmTc=OrTc and OmYy=OrYy and OmChr=OrChr and OmNo=OrNo and OmPrtKey=OrPrtKey)='N'      
    and @wSysDb='-'      
   If @Row> 0      
     Begin        
      select @errmsg= 'Cannot Delete Order Rm Of an UnAcknowledged Web Order'        
      goto ErrHndlr        
     end        
       
   /*****If  db= 'WEB' then Insert/Update  Not allowed in OrdMst  for OdChr<>'WEB'/OdChr='WEB' and OmSubmitYN='Y' ******/      
   /* Current Partition */       
   Select @Row=Count('x') From Deleted where OrPrtKey=@wCurrPrtn and @wSysDb='WEB' and       
   (OrChr<>'WEB' Or (OrChr='WEB' and (Select OmSubmitYN From OrdMst where OmCoCd=OrCoCd and OmTc=OrTc and       
        OmYy=OrYy and OmChr=OrChr and OmNo=OrNo and OmPrtKey=OrPrtKey)='Y')      
    )       
   If @Row> 0      
     Begin        
      select @errmsg= 'Cannot Delete  Ord Rm  from Orders Other than Chr=''WEB'' and Submitted=''N'' '        
      goto ErrHndlr        
     end        
  End        
 End        
      
  /******** Key Cannot Be Changed ********/       
  if @mode = 'C'      
  Begin      
  /* Current Partition */      
  select @Row=Count('x') from deleted del where del.OrPrtKey=@wCurrPrtn and      
    not exists (Select 'x' from inserted ins where ins.OrIdNo=del.OrIdNo and ins.OrCoCd= del.OrCoCd       
     and ins.OrTc= del.OrTc and ins.OrYy= del.OrYy and ins.OrChr= del.OrChr       
     and ins.OrNo= del.OrNo and ins.OrSr= del.OrSr and ins.OrSrNo= del.OrSrNo and Ins.OrPrtKey=del.OrPrtKey)      
  If @Row > 0      
  begin       
   Select @ErrMsg='Cannot Change Company Code / Tc/Yy/Chr/No/Sr/SrNo'      
   GoTo ErrHndlr      
  end      
 End      
      
 if @mode in ('A','C')      
 begin      
  /****** OrderSr should already exist in OrdDsg table ******/      
  /* Current Partition */      
  select @row=count('x') from inserted where OrPrtKey=@wCurrPrtn and      
   not exists (select 'x' from OrdDsg where OdIdNo=OrOdIdNo and OdCoCd= OrCoCd       
   and OdTc= OrTc and OdYy= OrYy and OdChr= OrChr       
   and OdNo= OrNo and OdSr= OrSr and OdPrtKey=OrPrtKey)      
  if @row>0       
  Begin      
   select @errmsg='Voucher Sr Not Defined In Order Design'      
   goto ErrHndlr      
  end      
      
  /******  OrSrNo  >  0  ******/      
  select @row=count('x') from inserted where OrPrtKey=@wCurrPrtn and OrSrNo<=0       
  if @row>0       
  Begin      
   select @errmsg='Raw Material Serial Number Cannot Be <= Zero'      
   goto ErrHndlr      
  end      
      
/* Cannot Delete JRI if exist in Txnd */
/*  select @row= count('x') from deleted del join OrdDsg on del.OrOdIdNo = OdIdNo and
    exists(Select 'x' from Txnd tx where tx.TdCoCd =OdCoCd and tx.TdBYy = OdDmStkYy and tx.TdBChr = OdDmStkChr and 
		tx.TdBNo = OdDmStkNo and tx.ModDt >= del.ModDt and tx.ModTime >= del.ModTime)       
  if @row> 0    
  begin    
   select @errmsg= 'Cannot Delete, Bag Transaction Exists for this Repair Bag'    
   goto ErrHndlr    
  end      
  */
  /* commented by urmi 28/02/04 **/      
  /******  OrWt  <>  0  ******/      
  /*select @row= count('x') from inserted where Round(OrWt, 3) = 0        
  if @row> 0       
  Begin      
   select @errmsg= 'Raw Material Weight Should Not Be = Zero'      
   goto ErrHndlr      
  end*/      
      
  /******  OrRmCtg  =  PMCd ('RMCTG')  ******/      
  /* Current Partition */      
  select @row= count('x') from inserted where OrPrtKey=@wCurrPrtn and      
   not exists (select 'x' from Param where PTyp= 'RMCTG' and PMCd= OrRmCtg)      
  if @row> 0       
  Begin      
   select @errmsg= 'Raw Material Category Not Defined'      
   goto ErrHndlr      
  end      
      
  /******  If OrRmCtg in('G', 'P', 'S') then OrLmeSal  >  0  else = 0 ******/        
  /* Current Partition */
  /* **** Manali 3.03 - 19/06/08 - 'L' included **** */      
  select @row= count('x') from inserted where OrPrtKey=@wCurrPrtn and      
   (OrRmCtg in('G', 'P', 'S', 'L') and Round(OrLmeSal, 2)<= 0) or       
   (OrRmCtg not in('G', 'P', 'S', 'L') and Round(OrLmeSal, 2)<> 0)       
  if @row> 0       
  begin      
   select @errmsg= 'LME Sale Value Should Be > Zero For Ctgs ''G'', ''P'', ''S'' and ''L'' and Zero For Other Ctgs'      
   goto ErrHndlr      
  end      
      
  /******  OrRmSCtg  =  PSCd ('RMSCTG', OrRmCtg)  ******/      
  /* Current Partition */      
  select @row=count('x') from inserted where OrPrtKey=@wCurrPrtn and    not exists (select 'x' from Param where PTyp='RMSCTG'       
   and PMCd=OrRmCtg and PSCd=OrRmSCtg)      
  if @row>0       
  Begin      
   select @errmsg='Raw Material Sub Category Not Defined'      
   goto ErrHndlr      
  end      
      
/*  ######################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ######################  */      
  /******  OrRmCd Not Blank OrRmCd  =  RmCd from RmMst (RmCtg= OrRmCtg, RmSCtg= OrRmSCtg)  ******/      
  /* Current Partition */       
  select @row= count('x') from Inserted Ins where OrPrtKey=@wCurrPrtn and OrRmCd = ''      
  if @row> 0      
  Begin      
   select @errmsg= 'Raw Material Code Can Not Blank'      
   goto ErrHndlr      
  end      
      
  /* Current Partition */      
  /**** Manali 3.8.0 - Invalid RmCd Allowed in Return Memo Order - Join with OrdMst added ****/
  select @row= count('x') from Inserted Ins 
		Join OrdMst On OmCoCd=OrCoCd and OmTc=OrTc and OmYy=OrYy and OmChr=OrChr and OmNo=OrNo and OmPrtKey=OrPrtKey
		where Ins.OrPrtKEy=@wCurrPRtn and OmReNo=0 and     
   Ins.OrRmCd<> IsNull((Select Del.OrRmCd From Deleted Del Where      
    /* Ins.OrIdNo=Del.OrIdNo and */ Ins.OrCoCd= Del.OrCoCd And Ins.OrTc = Del.OrTc And Ins.OrYy = Del.OrYy And      
    Ins.OrChr= Del.OrChr And Ins.OrNo = Del.OrNo And      
    Ins.OrSr= Del.OrSr And Ins.OrSrNo = Del.OrSrNo And Ins.OrPrtKey=Del.OrPrtKey), '') And      
   not exists (select 'x' from RmMst where RmCtg=OrRmCtg      
   and RmSCtg=OrRmSCtg and RmCd=OrRmCd and RmPrtKey=OrPrtKey And RmValidYN='Y')      
  if @row>0      
  Begin      
   select @errmsg='Raw Material Code Not Defined/ Invalid'      
   goto ErrHndlr      
  end      
  
  /* Current Partition */      
  select @row=count('x') from inserted where OrPrtKey=@wCurrPrtn and      
   not exists (select 'x' from RmMst where RmCtg= OrRmCtg      
   and RmSCtg= OrRmSCtg and RmCd= OrRmCd and RmPrtKey=OrPrtKey)      
  if @row>0      
  Begin      
   select @errmsg='Raw Material Code Not Defined'      
   goto ErrHndlr      
  end      
/*  #########################################  Manoj  #########################################  */      
      
      
/* temporary stopped by rn      
  select @row=count('x') from inserted where (OrRmCtg not in ('D','C') and (OrSetSCd<>''  or OrSetSalRt<>0 or OrSetSalVal<>0      
         or OrSetCstRt<>0 or OrSetCstVal<>0 or OrWsQty<>0 or OrHsQty<>0 or OrSubShp<>''))       
                    or (OrRmCtg not in ('P','G','S') and (OrAlyCd<>'' or OrAlySalRt<>0 or OrAlyCstRt<>0))      
  if @row>0       
  Begin      
   select @errmsg='Setting Cd,Rt & Val should be entered only for Dia or ColStn.  Aly Cd,Rt & Val for Gold or Platinum'      
   goto ErrHndlr      
  end      
*/      
      
/*  ######################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ######################  */      
  /******  OrSetSCd <> Blank,  for Raw Material Categories Dias & Col Stns ******/      
  /* Current Partition */      
  select @row= count('x') from inserted where OrPrtKey=@wCurrPrtn and OrRmCtg in ('C', 'D') and OrSetSCd= ''      
  if @row> 0      
  Begin      
   select @errmsg= 'Setting Code Cannot Be Blank For This Rm Code'      
   goto ErrHndlr      
  end      
     
  /******  OrSetSCd  =  PSCd ('LABSCD', 'SET')  for Dias & Col Stns ******/      
  /* Current Partition */      
  select @row= count('x') from inserted Ins where OrPrtKey=@wCurrPrtn and (OrRmCtg in ('D', 'C') or OrSetSCd<> '')      
   And Ins.OrSetSCd<> IsNull((Select Del.OrSetSCd From Deleted Del Where      
    Ins.OrIdNo=Del.OrIdNo and Ins.OrCoCd = Del.OrCoCd And Ins.OrTc = Del.OrTc And Ins.OrYy = Del.OrYy And      
    Ins.OrChr = Del.OrChr And Ins.OrNo = Del.OrNo And      
    Ins.OrSr = Del.OrSr And Ins.OrSrNo = Del.OrSrNo And Ins.OrPrtKey=Del.OrPRtKey), '')      
   and not exists (select 'x' from Param where PTyp= 'LABSCD'      
   and PMCd='SET' and PSCd= OrSetSCd And PValidYN = 'Y')      
  if @row> 0      
  Begin      
   select @errmsg= 'Setting Code Not Defined/ Invalid'      
   goto ErrHndlr      
  end      
      
/* (Bef 2.06)      
  select @row=count('x') from inserted where (OrRmCtg in ('D', 'C') or OrSetSCd<> '')      
   and not exists (select 'x' from Param where PTyp= 'LABSCD'      
   and PMCd='SET' and PSCd= OrSetSCd)      
  if @row>0      
  Begin      
   select @errmsg='Setting Code Not Defined'      
   goto ErrHndlr      
  end      
*/      
/* #########################################  Manoj  #########################################  */      
      
  /******  OrSetSCd = Blank,  OrSetSalRt = 0,  OrSetSalVal = 0  for Raw Material Categories Gold, Platinum and Metal ******/      
  /* Current Partition */      
  --5717 OrSubShp <> '' removed. SunShp can be enter for GPSL also
  select @row=count('x') from inserted where OrPrtKey=@wCurrPrtn and OrRmCtg Not in ('C', 'D') and (OrSetSCd<> '' or OrSetSalRt<> 0 or OrSetSalVal<> 0       
    or OrSetCstRt<> 0 or OrSetCstVal<> 0 )     if @row>0       
  Begin      
   select @errmsg='Setting Code/ Rate/ Value  Should Not Be Entered For This Rm Code'      
   goto ErrHndlr      
  end      
      
  /******  OrQty = OrWsQty + OrHsQty  for Dias & Col Stns ******/      
  /* Current Partition */      
  select @row=count('x') from inserted where OrPrtKey=@wCurrPrtn and OrRmCtg in ('D','C')  and OrQty <> OrWsQty + OrHsQty       
  if @row>0       
  Begin      
   select @errmsg= 'Rm Qty Should Be Equal To  Wax Set Qty + Hand Set Qty'      
   goto ErrHndlr      
  end      
      
  /******  OrSubShp  =  PMCd ('SUBSHP')  ******/      
  /* Current Partition */      
  select @row=count('x') from inserted where OrPrtKEy=@wCurrPrtn and OrSubShp<> '' and       
   not exists (select 'x' from Param where PTyp='SUBSHP' and PMCd=OrSubShp)      
  if @row>0       
  Begin      
   select @errmsg='Sub Shape Not Defined'      
   goto ErrHndlr      
  end      
      
  /******  OrAlyCd  =  Blank  or  RmCd from RmMst (RmCtg= 'A')   when  OrRmCtg= 'G' or 'P' or 'S'  ******/      
  /* Current Partition */
  /* **** Manali 3.03 - 19/06/08 - 'L' included **** */      
  select @row=count('x') from inserted Where OrPrtKey=@wCurrPRtn and OrAlyCd<> ''     
  and OrRmCtg in ('G', 'P', 'S', 'L')     and not exists (select 'x' from RmMst where RmCtg= 'A' and RmCd= OrAlyCd and RmPRtKey=OrPrtKey)      
  if @row>0       
  Begin      
   select @errmsg='Alloy Code Not Defined '+ (select max(OrAlyCd) from inserted )    
   goto ErrHndlr      
  end      
      
  /******  OrMainMet  =  PMCd ('YN')  ******/      
  /* Current Partition */      
  select @row=count('x') from inserted where OrPRtKey=@wCurrPrtn and     not exists (select 'x' from Param where PTyp='YN' and PMCd=OrMainMet)      
  if @row>0       
  Begin      
   select @errmsg='Enter Main Metal (Y/N)'      
   goto ErrHndlr      
  end      

  /****** Sachin 3.3.0 (30-06-08) - [28. Valid OrLn1/OrLn2 Check]   ******/      	
  Select @row=Count('x') From Inserted 			
		Where @mode In ('A','C') and (update(OrLn1) Or Update(OrLn2) Or Update(OrRmCd)) And OrPRtKey=@wCurrPrtn And OrRmCtg In ('D','C') And
		Exists(Select 'x' From Head Where HCoCd= OrCoCd and HCd= @wHCd And hChkVldRmLnYN='Y') And
		Not Exists(Select 'x' From VldRmLn Where VrsRmCtg=OrRmCtg 
						And VrsRmId=(Select RmId From RmMst Where RmCtg=OrRmCtg And RmCd=OrRmCd) And VrsLn1=OrLn1 And VrsLn2=OrLn2)		 
  if @row>0       
  Begin      
   select @errmsg='Ln1 Not Defined/Ln1 And Ln2 Combination Not Defined In RM Ln Master'      
   goto ErrHndlr      
  end      
	
  /******  If HShowPrdRmDets= 'N' then OrWt = OrPrdWt ******/      
  /* Current Partition */      
  select @row= count('x') from inserted where OrPRtKey=@wCurrPrtn and      
   (Str(OrWt, 16, 3)<> Str(OrPrdWt, 16, 3)       
    or OrQty<> OrPrdQty)       
   and (Select HShowPrdRmDets from Head where HCoCd= OrCoCd and HCd= @wHCd) = 'N'    if @row> 0       
  Begin      
   select @errmsg= 'Rm Wt/ Qty Should Be Equal To Rm Prd Wt/ Qty As Per The Option Selected In Head File'      
   goto ErrHndlr      
  end      
      
  /****** Only One Main Metal Record Is allowed Per OrdSr ******/      
  /* Current Partition */      
  select @row=count('x') from inserted ins where OrPrtKey=@wCurrPrtn and      
   (select count('x') from OrdRm Orm where ins.OrOdIdNo=Orm.OrOdIdNo and  ins.OrCoCd= Orm.OrCoCd and       
   ins.OrTc= Orm.OrTc and ins.OrYy= Orm.OrYy and ins.OrChr= Orm.OrChr and       
   ins.OrNo= Orm.OrNo and ins.OrSr= Orm.OrSr and Ins.OrPrtKEy=Orm.OrPrtKey and Orm.OrMainMet = 'Y') > 1       
  if @row>0       
  Begin      
   select @errmsg='Cannot Have More Than One Main Metal'      
   goto ErrHndlr      
  end      
      
/*  select @row=count('x') from deleted del where not exists       
   (Select 'x' from Inserted Ins where ins.OrTc= del.OrTc and ins.OrYy= del.OrYy  and ins.OrChr= del.OrChr and ins.OrNo= del.OrNo       
    and ins.OrSr= del.OrSr and ins.OrSrNo= del.OrSrNo and ins.OrRmCtg=del.OrRmCtg and  ins.OrRmSCtg= del.OrRmSCtg and       
           ins.OrRmCd= del.OrRmCd)  and       
   (Select OdExpQty from OrdDsg where OdTc= del.OdTc and OdYy= del.OdYy       
      and OdChr= del.OrChr and OdNo= del.OrNo and OdSr= del.OrSr )>0 or      
   (exists (Select 'x' from Bag where  BOdTc= ins.OrTc and BOdYy= ins.OrYy       
      and BOdChr= ins.OrChr and BOdNo= ins.OrNo and BOdSr= ins.OrSr and      
    isnull((select sum().... ),0)))      
      
      
  if @row>0       Begin      
   select @errmsg='Cannot Change Raw Material Category/Sub Category/Code'      
   goto ErrHndlr      
  end      
*/      
   /****** Cannot change Rm Ctg/ SCtg / Code if Export has taken place or       
          Transaction of the Rm Code has taken place or       
          Fg rec (without bag) exists for the Rm Code ******/      
  /************ The best check would have been if you could check whether the Rm Code which is changed or deleted is currently in the       
         bag or not (do not do it now) ************/      
  /* Complete Partition - Check */     
  /****Sangeeta **Commented ExpOrder check (or condition)*****************************/ 
  select @row= count('x') from deleted del join inserted ins on ins.OrIdNo=del.OrIdNo and ins.OrCoCd= del.OrCoCd and       
   ins.OrTc= del.OrTc and ins.OrYy= del.OrYy and ins.OrChr= del.OrChr and       
   ins.OrNo= del.OrNo and ins.OrSr= del.OrSr and ins.OrSrNo= del.OrSrNo 
   where (ins.OrRmCtg<> del.OrRmCtg or ins.OrRmSCtg<> del.OrRmSCtg or ins.OrRmCd<> del.OrRmCd       
    or ins.OrMainMet<> del.OrMainMet)        
   And       
   ( (Select OdExpQty from OrdDsg where OdIdNo=del.OrOdIdNo and OdCoCd= del.OrCoCd and OdTc= del.OrTc and OdYy= del.OrYy       
      and OdChr= del.OrChr and OdNo= del.OrNo and OdSr= del.OrSr and OdPrtKey=del.OrPRtKey )> 0       
    or      
    (exists (Select 'x' from Bag where BOdIdNo=del.OrOdIdNo and BCoCd= del.OrCoCd and BOdTc= del.OrTc and BOdYy= del.OrYy       
      and BOdChr= del.OrChr and BOdNo= del.OrNo and BOdSr= del.OrSr and       
      isnull((select count(TdTc) from Txnd where TdBIdNo=BIdNo and TdCoCd= BCoCd       
       and TdBYy= BYy and TdBChr= BChr and TdBNo= BNo and TDRmCd= del.OrRmCd),0) > 0))       
  /*  or      
    (exists (Select 'x' from Fgd where FdCoCd= del.OrCoCd and FdExpOdTc= del.OrTc and FdExpOdYy= del.OrYy       
      and FdExpOdChr= del.OrChr and FdExpOdNo= del.OrNo and FdExpOdSr= del.OrSr and      
      isnull((select count(FrTc) from FgRm where FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc       
      and FrYy= FdYy and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr and FrRmCd= del.OrRmCd), 0) > 0 )) */      
   )      
  if @row> 0       
  Begin      
   select @errmsg='Cannot Change Raw Material Category/Sub Category/Code/MainMetal'      
   goto ErrHndlr      
  end      
      
/*      
   And not exists(Select 'x' from OrdRm Orm where Orm.OrTc= del.OrTc       
     and Orm.OrYy= del.OrYy and Orm.OrChr= del.OrChr       
     and Orm.OrNo= del.OrNo and Orm.OrSr= del.OrSr       
     and Orm.OrRmCtg= del.OrRmCtg and Orm.OrRmSCtg= del.OrRmSCtg       
     and Orm.OrRmCd= del.OrRmCd and Orm.OrMainMet= del.OrMainMet)       
      
*/    
      
/*  select @row=count('x') from deleted del join inserted ins on ins.OrTc= del.OrTc and ins.OrYy= del.OrYy and ins.OrChr= del.OrChr and       
   ins.OrNo= del.OrNo and ins.OrSr= del.OrSr and ins.OrSrNo= del.OrSrNo       
   where (ins.OrRmCtg<> del.OrRmCtg or ins.OrRmSCtg<> del.OrRmSCtg or ins.OrRmCd<> del.OrRmCd       
    or ins.OrMainMet<> del.OrMainMet)  and       
   ( (Select OdExpQty from OrdDsg where OdTc= del.OrTc and OdYy= del.OrYy       
      and OdChr= del.OrChr and OdNo= del.OrNo and OdSr= del.OrSr )>0       
    or      
    (exists (Select 'x' from Bag where  BOdTc= del.OrTc and BOdYy= del.OrYy       
      and BOdChr= del.OrChr and BOdNo= del.OrNo and BOdSr= del.OrSr and      
      isnull((select count(TdTc) from Txnd where TDRmCd= del.OrRmCd and TdBYy= BYy       
       and TdBChr= BChr and TdBNo= BNo),0) > 0 ))      
    or      
  (exists (Select 'x' from Fgd where FdExpOdTc= del.OrTc and FdExpOdYy= del.OrYy       
      and FdExpOdChr= del.OrChr and FdExpOdNo= del.OrNo and FdExpOdSr= del.OrSr and      
      isnull((select count(FrTc) from FgRm where FrRmCd= del.OrRmCd and FrTc= FdTc and FrYy= FdYy       
       and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr), 0) > 0 ))      
   )      
  if @row>0       
  Begin      
   select @errmsg='Cannot Change Raw Material Category/Sub Category/Code/MainMetal'      
   goto ErrHndlr      
  end      
*/      
      
 end      
      
 if @mode in ('D')      
 Begin      
  /****** Cannot delete record if Export has taken place or       
         Transaction of the Rm Code has taken place or       
         Finished Goods rec (without bag) exists for the Rm Code ******/      
  /* Complete Partition - Check */       
  select @row= count('x') from deleted del       
   where del.OrPrtKey=@wCurrPrtn and      
   ( (Select OdExpQty from OrdDsg where OdIdNo=del.OrOdIdNo and OdCoCd= del.OrCoCd and OdTc= del.OrTc and OdYy= del.OrYy       
     and OdChr= del.OrChr and OdNo= del.OrNo and OdSr= del.OrSr and OdPrtKey=del.OrPrtKey)> 0       
    or       
    (exists (Select 'x' from Bag where BOdIdNo=del.OrOdIdNo and BCoCd= del.OrCoCd and BOdTc= del.OrTc and BOdYy= del.OrYy       
     and BOdChr= del.OrChr and BOdNo= del.OrNo and BOdSr= del.OrSr and BPrtKey=del.OrPrtKey and      
     isnull((select count(TdTc) from Txnd where TdBIdNo=BIdNo and TdCoCd= BCoCd and TdBYy= BYy       
      and TdBChr= BChr and TdBNo= BNo and TdRmCd= del.OrRmCd and TdPrtKey=del.OrPrtKey), 0) > 0 ))      
    or      
    (exists (Select 'x' from Fgd where FdCoCd= del.OrCoCd and FdExpOdTc= del.OrTc and FdExpOdYy= del.OrYy       
     and FdExpOdChr= del.OrChr and FdExpOdNo= del.OrNo and FdExpOdSr= del.OrSr and FdPrtKey=del.OrPrtKey and       
     isnull((select count(FrTc) from FgRm where FrFdIdNo=FdIdNo and FrCoCd= FdCoCd and FrTc= FdTc and FrYy= FdYy       
      and FrChr= FdChr and FrNo= FdNo and FrSr= FdSr and FrRmCd= del.OrRmCd and FrPrtKey=del.OrPrtKey), 0) > 0 ))      
   )      
  if @row> 0      
  Begin      
   select @errmsg= 'Cannot Delete As Transaction Exists For Rm Code'      
   goto ErrHndlr      
  end      
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

ALTER Procedure [dbo].[InsOrderCpy]
(
	@wTargetOdCoCd varchar(3),
	@wTargetOdTc varchar(3),
	@wTargetOdYy varchar(2),
	@wTargetOdChr varchar(3),
	@wTargetOdNo int,
	@wOdSr int,
	@wOdDmCd varchar(15),
	@wOdSfx varchar(15),
	@wOdDmSz varchar(5),
	@wOdPrdSeq varchar(80),
	@wOdDelDt date,
	@wOdOrdEnt float,
	@wOdOrdQty float,
	@wOdPrdEnt float,
	@wOdPrdQty float,
	@wOdCalcPrc float,
	@wOdSalPrc float,
	@wOdCstPrc float,
	@wOdDmPrdInst varchar(250),
	@wOdCmPrdInst varchar(250),
	@wOdCmStmpInst varchar (90),
	@wOdSzInst varchar(90),
	@wOdPrtCd varchar(2),
	@wOdVaCtg varchar(5),
	@wOdKt varchar(5),
	@wOdMulBy float,
	@wOdFixPrc varchar(1),
	@wOdGldAs varchar(5),
	@wOdIWtEqOrd varchar(1),
	@wOdIWtFrOrd varchar(5),
	@wOdIGldRtEqOrd varchar(1),
	@wOdBagPcs float,
	@wOdDmCol varchar(5),
	@wOdGldAsWt float,
	@wOdExpDelDt date,
	@wOdSfxDesc varchar(255),
	@wOdWh varchar(2),
	@wOdByy varchar(2),
	@wOdBChr varchar(3),
	@wOdBNo int,
	@wOdMrp float,
	@wOdMrpDisc float,
	@wOdLine varchar(5),
	@wOdPoNo varchar(20),
	@wOdSubCust varchar(8),
	@ws_baseVch varchar(50),
	@wOdCell  varchar(8),
	@wOdDmStkCoCd varchar(3),
	@wOdDmStkYy varchar(2),
	@wOdDmStkChr varchar(3),
	@wOdDmStkNo int,
	@wOdChgPtrOnStwYN varchar(1),
	@wOdMrpMulBy float,
	@wOdILabWtFrOrd varchar(1),
	@wOdInvAllBags varchar(1),
	@wOdLabAs varchar(5),
	@wOdLabAsWt float,
	@wOdSalRem varchar(100),
	@wOdOmCtCd varchar(8),
	@wOdCrmFixPrcYN varchar(1),
	@wOdPicNm varchar(45),
	@wOdGmChk varchar(8),
	@wOdModUsr varchar(5),
	@wOdGrMet varchar(8),
	@wOdGrDia varchar(8),
    @wOdGrCS varchar(8),
    @wOdGrFin varchar(8),
	@wOdMinWt float,
	@wOdMaxWt float,
	@wOdChkTol varchar(1),
	@wOdSubRem varchar(120),
	@wOdDmIdNo int
    ,@wOdMinDiaTolWt float 
    ,@wOdMaxDiaTolWt float
	,@wOdBaseIdKey int
	,@wOdRefIdKey int
	,@wOdCls char(1)
	,@wOdHld char(1)
	,@wOdWDiaAvlblDt date
	,@wOdHDiaAvlblDt date
	,@wOdFndAvlblDt date
	,@wOdInsDt date
	,@wOdOmDt date 
	,@wOdDmTcTyp varchar(2)
	,@wOdOmIdNo int
	,@wOdOmCmCd varchar(8)
	,@qOdPerPcDiaWt float
	,@qOdPerPcCsWt float
	,@qOdPerPcDiaQty int
	,@qOdPerPcCsQty int
	,@qOdMainMetRmCtg varchar(1)
	,@qOdCstLabAsWt float
	,@qOdPerPcDiaPrdWt float
	,@qOdPerPcCsPrdWt float
	,@qMode int
	,@wOdRepGrWt float
)
--OdSr - RedSr
as
Begin

	if @qMode=1 
	begin

		INSERT INTO  dbo . OrdDsg 
           ( OdTc 
           , OdYy 
           , OdChr 
           , OdNo 
           , OdSr 
           , OdDmCd 
           , OdSfx 
           , OdDmSz 
           , OdPrdSeq 
           , OdDelDt 
           , OdOrdEnt 
           , OdOrdQty 
           , OdExpQty 
           , OdPrdEnt 
           , OdPrdQty 
           , OdFgQty 
           , OdCalcPrc 
           , OdSalPrc 
           , OdCstPrc 
           , OdDmPrdInst 
           , OdCmPrdInst 
           , OdCmStmpInst 
           , OdSzInst 
           , OdPrtCd 
           , OdHld 
           , OdHldDesc 
           , OdVaCtg 
           , OdKt 
           , OdMulby 
           , OdFixPrc 
           , OdGldAs 
           , OdIWtEqOrd 
           , OdIWtFrOrd 
           , OdIGldRtEqOrd 
           , OdCls 
           , OdBagPcs 
           , OdDmCol 
           , OdGldAsWt 
           , ModUsr 
           , ModDt 
           , ModTime 
           , OdExpDelDt 
           , OdSfxDesc 
           , OdCoCd 
           , OdWh 
           , OdWDiaAvlblDt 
           , OdHDiaAvlblDt 

           , OdLine 
           , OdSalRem 
           , OdOmCtCd 
           , OdMinWt 
           , OdMaxWt 
           , OdBYy 
           , OdBChr 
           , OdBNo 
           , OdGmChk 
           , OdChkTol 
           , OdGrMet 
           , OdGrDia 
           , OdGrCS 
           , OdGrFin 
           , OdCrmFixPrcYN 
           , OdSubRem 
           , OdLabAs 
           , OdLabAsWt 
           , OdOmCmCd 
           , OdOmDt 
           , OdOmIdNo 
		   ,OdDmIdNo
           , OdPrtKey 
           , OdPrtFgQty 
           , OdPrtExpQty 
           , OdPicNm 
           , OdPoNo 
           , OdILabWtFrOrd 
           , OdChgPtrOnStwYN 
           , InsDt 
           , OdMinDiaTolWt 
           , OdMaxDiaTolWt 
           , OdBaseIdKey 
           , OdRefIdKey 
           , OdMrpMulby 
           , OdMrp 
           , OdDmStkYy 
           , OdDmStkChr 
           , OdDmStkNo 
			, OdJLRmCd 
			, OdJLLn1 
			, OdJLLn2 
			, OdJLLn3 
			, OdJLLotNo 
			, OdJLQty 
			, OdJLWt 
			, OdJLExpWt 
			, OdJLVchRt 
			, OdJLRateByQW 
			, OdJLVchVal 
			, OdInvAllBags 
			,OdCell
			,OdRepGrWt
			,OdDmStkCoCd
			,OdMrpDisc
			,OdFndAvlblDt 
			,OdDmTcTyp
			,OdSubCust 
           )
		SELECT     
		   @wTargetOdTc,
           @wTargetOdYy,
           @wTargetOdChr,
           @wTargetOdNo,
		   @wOdSr,
           @wOdDmCd,
           @wOdSfx, 
           @wOdDmSz, 
		   @wOdPrdSeq,
           @wOdDeldt, 
           @wOdOrdEnt, 
           @wOdOrdQty, 
		   0 as OdExpQty,		   
		   @wOdPrdEnt,
		   @wOdPrdQty, 
		   0 as OdFgQty,
           @wOdCalcPrc,
           @wOdSalPrc,
			@wOdCstPrc,
           @wOdDmPrdInst,
           @wOdCmPrdInst,
           @wOdCmStmpInst,
		   @wOdSzInst,
           @wOdPrtCd,
           @wOdHld,
           '' as OdHldDesc,
           @wOdVaCtg
           ,@wOdKt
           ,@wOdMulBy 
           ,@wOdFixPrc
		   ,@wOdGldAs
           ,@wOdIWtEqOrd
           ,@wOdIWtFrOrd
           ,@wOdIGldRtEqOrd
		   ,@wOdCls
           ,@wOdBagPcs
           ,@wOdDmCol
           ,@wOdGldAsWt
          ,@wOdModUsr
			,convert(varchar(8),getdate(),3)
			,datepart(hh,getdate())+(convert(float,datepart(mi,getdate()))/100)
           ,@wOdExpDelDt
		   ,@wOdSfxDesc
		   ,@wTargetOdCoCd
           ,@wOdWh
           ,@wOdWDiaAvlblDt 
           ,@wOdHDiaAvlblDt 
           ,@wOdLine
           ,@wOdSalRem
           ,@wOdOmCtCd
           ,@wOdMinWt
           ,@wOdMaxWt
           ,@wOdBYy
           ,@wOdBChr
           ,@wOdBNo
           ,@wOdGmChk
           ,@wOdChkTol
           ,@wOdGrMet
           ,@wOdGrDia
           ,@wOdGrCS
           ,@wOdGrFin
           ,@wOdCrmFixPrcYN
           ,@wOdSubRem
           ,@wOdLabAs
           ,@wOdLabAsWt
           ,@wOdOmCmCd
           ,@wOdOmDt
           ,@wOdOmIdNo
           ,@wOdDmIdNo
           ,'C'
           ,0 as OdPrtFgQty
           ,0 as OdPrtExpQty
           ,@wOdPicNm
           ,@wOdPoNo
           ,@wOdILabWtFrOrd
		   ,@wOdChgPtrOnStwYN
           ,@wOdInsDt/*convert(varchar(8),getdate(),3) as InsDt*/
           ,@wOdMinDiaTolWt
           ,@wOdMaxDiaTolWt
		   ,@wOdBaseIdKey
		   ,@wOdRefIdKey 
           ,@wOdMrpMulby
           ,@wOdMrp
           ,@wOdDmStkYy 
           ,@wOdDmStkChr
           ,@wOdDmStkNo
		    ,'' as OdJLRmCd
      ,0 as OdJLLn1
      ,0 as OdJLLn2
      ,0 as OdJLLn3
      ,'' as OdJLLotNo
      ,0 as OdJLQty
      ,0.0 as OdJLWt
      ,0.0 as OdJLExpWt
      ,0.0 as OdJLVchRt
      ,'' OdJLRateByQW
      ,0.0 as OdJLVchVal
      ,@wOdInvAllBags
	  ,@wOdCell
	  ,@wOdRepGrWt
	  ,@wOdDmStkCoCd 
	  ,@wOdMrpDisc
      ,@wOdFndAvlblDt 
	  ,@wOdDmTcTyp 
	  ,@wOdSubCust 
	  end
	  else
	  begin
		UPDATE OrdDsg set 
           OdDmcd=@wOdDmCd,
           OdSfx=@wOdSfx, 
           OdDmSz=@wOdDmSz, 
		   OdPrdSeq=@wOdPrdSeq,
           OdDeldt=@wOdDelDt, 
           OdOrdEnt=@wOdOrdEnt,
           OdOrdQty=@wodordqty,
		   OdExpQty=0,		   
		   OdPrdEnt=@wOdPrdEnt,
		   OdPrdQty=@wOdPrdEnt, 
		   OdFgQty=0,
           OdCalcPrc=@wOdCalcPrc,
           OdSalPrc=@wOdSalPrc,
			OdCstPrc=@wOdCstPrc,
           OdDmPrdInst=@wOdDmPrdInst,
           OdCmPrdInst=@wOdCmPrdInst,
           OdCmStmpInst=@wOdCmStmpInst,
		   OdSzInst=@wOdSzInst,
           OdPrtCd=@wOdPrtCd,
           OdHld=@wOdHld,
           OdHldDesc='',
           OdVaCtg=@wOdVaCtg,
           OdKt=@wOdKt,
           OdMulBy=@wOdMulBy ,
           OdFixPrc=@wOdFixPrc,
		   OdGldAs=@wOdGldAs,
           OdIWtEqOrd=@wOdIWtEqOrd,
           OdIWtFrOrd=@wOdIWtFrOrd,
           OdIGldRtEqOrd=@wOdIGldRtEqOrd,
		   OdCls=@wOdCls,
           OdBagPcs=@wOdBagPcs,
           OdDmCol=@wOdDmCol,
           OdGldAsWt=@wOdGldAsWt,
          ModUsr=@wOdModUsr,
           OdExpDelDt=@wOdExpDelDt ,
		   OdSfxDesc=@wOdSfxDesc,
		
           OdWh=@wOdWh,
           OdWDiaAvlblDt =@wOdWDiaAvlblDt ,
           OdHDiaAvlblDt =@wOdHDiaAvlblDt ,
           OdLine=@wOdLine,
           OdSalRem=@wOdSalRem,
           OdOmCtCd=@wOdOmCtCd,
           OdMinWt=@wOdMinWt,
           OdMaxWt=@wOdMaxWt,
           OdBYy=@wOdBYy,
           OdBChr=@wOdBChr,
           OdBNo=@wOdBNo,
           odGmChk=@wOdGmChk,
           OdChkTol=@wOdChkTol,
           OdGrMet=@wOdGrMet,
           OdGrDia=@wOdGrDia,
           OdGrCS=@wOdGrCS,
           OdGrFin=@wOdGrFin,
           OdCrmFixPrcYN=@wOdCrmFixPrcYN,
          OdSubRem=@wOdSubRem,
           OdLabAs=@wOdLabAs,
           OdLabAsWt=@wOdLabAsWt,
           OdOmCmCd=@wOdOmCmCd,
           OdOmDt=@wOdOmDt,
           OdPicNm=@wOdPicNm,
           OdPoNo=@wOdPoNo,
           OdILabWtFrOrd=@wOdILabWtFrOrd,
		   OdChgPtrOnStwYN=@wOdChgPtrOnStwYN,
      
           OdMinDiaTolWt=@wOdMinDiaTolWt,
           OdMaxDiaTolWt=@wOdMaxDiaTolWt,
		   OdBaseIdKey=@wOdBaseIdKey,
		   OdRefIdKey =@wOdRefIdKey ,
           OdMrpMulby=@wOdMrpMulby,
          OdMrp=@wOdMrp,
           OdDmStkYy=@wOdDmStkYy ,
           OdDmStkChr=@wOdDmStkChr,
           OdDmStkNo=@wOdDmStkNo,
      OdInvAllBags=@wOdInvAllBags,
	  OdCell=@wOdCell,
		OdRepGrWt = @wOdRepGrWt,
	  OdDmStkCoCd =@wOdDmStkCoCd ,
	  OdMrpDisc=@wOdMrpDisc,
      OdFndAvlblDt=@wOdFndAvlblDt ,
	  OdDmTcTyp =@wOdDmTcTyp ,
	  OdSubCust =@wOdSubCust 
		where
		          OdCoCd=@wTargetOdCoCd and OdTc=@wTargetOdTc and  OdYy =@wTargetOdYy and OdChr =@wTargetOdChr and OdNo=@wTargetOdNo and OdSr=@wOdSr 

	  end

/*	return select @@IDENTITY */

end
GOGO

ALTER Procedure [dbo].[InsOrderCpy]
(
	@wTargetOdCoCd varchar(3),
	@wTargetOdTc varchar(3),
	@wTargetOdYy varchar(2),
	@wTargetOdChr varchar(3),
	@wTargetOdNo int,
	@wOdSr int,
	@wOdDmCd varchar(15),
	@wOdSfx varchar(15),
	@wOdDmSz varchar(5),
	@wOdPrdSeq varchar(80),
	@wOdDelDt date,
	@wOdOrdEnt float,
	@wOdOrdQty float,
	@wOdPrdEnt float,
	@wOdPrdQty float,
	@wOdCalcPrc float,
	@wOdSalPrc float,
	@wOdCstPrc float,
	@wOdDmPrdInst varchar(250),
	@wOdCmPrdInst varchar(250),
	@wOdCmStmpInst varchar (90),
	@wOdSzInst varchar(90),
	@wOdPrtCd varchar(2),
	@wOdVaCtg varchar(5),
	@wOdKt varchar(5),
	@wOdMulBy float,
	@wOdFixPrc varchar(1),
	@wOdGldAs varchar(5),
	@wOdIWtEqOrd varchar(1),
	@wOdIWtFrOrd varchar(5),
	@wOdIGldRtEqOrd varchar(1),
	@wOdBagPcs float,
	@wOdDmCol varchar(5),
	@wOdGldAsWt float,
	@wOdExpDelDt date,
	@wOdSfxDesc varchar(255),
	@wOdWh varchar(2),
	@wOdByy varchar(2),
	@wOdBChr varchar(3),
	@wOdBNo int,
	@wOdMrp float,
	@wOdMrpDisc float,
	@wOdLine varchar(5),
	@wOdPoNo varchar(20),
	@wOdSubCust varchar(8),
	@ws_baseVch varchar(50),
	@wOdCell  varchar(8),
	@wOdDmStkCoCd varchar(3),
	@wOdDmStkYy varchar(2),
	@wOdDmStkChr varchar(3),
	@wOdDmStkNo int,
	@wOdChgPtrOnStwYN varchar(1),
	@wOdMrpMulBy float,
	@wOdILabWtFrOrd varchar(1),
	@wOdInvAllBags varchar(1),
	@wOdLabAs varchar(5),
	@wOdLabAsWt float,
	@wOdSalRem varchar(100),
	@wOdOmCtCd varchar(8),
	@wOdCrmFixPrcYN varchar(1),
	@wOdPicNm varchar(45),
	@wOdGmChk varchar(8),
	@wOdModUsr varchar(5),
	@wOdGrMet varchar(8),
	@wOdGrDia varchar(8),
    @wOdGrCS varchar(8),
    @wOdGrFin varchar(8),
	@wOdMinWt float,
	@wOdMaxWt float,
	@wOdChkTol varchar(1),
	@wOdSubRem varchar(120),
	@wOdDmIdNo int
    ,@wOdMinDiaTolWt float 
    ,@wOdMaxDiaTolWt float
	,@wOdBaseIdKey int
	,@wOdRefIdKey int
	,@wOdCls char(1)
	,@wOdHld char(1)
	,@wOdWDiaAvlblDt date
	,@wOdHDiaAvlblDt date
	,@wOdFndAvlblDt date
	,@wOdInsDt date
	,@wOdOmDt date 
	,@wOdDmTcTyp varchar(2)
	,@wOdOmIdNo int
	,@wOdOmCmCd varchar(8)
	,@qOdPerPcDiaWt float
	,@qOdPerPcCsWt float
	,@qOdPerPcDiaQty int
	,@qOdPerPcCsQty int
	,@qOdMainMetRmCtg varchar(1)
	,@qOdCstLabAsWt float
	,@qOdPerPcDiaPrdWt float
	,@qOdPerPcCsPrdWt float
	,@qMode int
	,@wOdRepGrWt float
	,@wOdMasterPrc float
	,@wOdTag float
	,@wOdProfit float
	,@wOdProfitPer float 
)
--OdSr - RedSr
as
Begin

	if @qMode=1 
	begin

		INSERT INTO  dbo . OrdDsg 
           ( OdTc 
           , OdYy 
           , OdChr 
           , OdNo 
           , OdSr 
           , OdDmCd 
           , OdSfx 
           , OdDmSz 
           , OdPrdSeq 
           , OdDelDt 
           , OdOrdEnt 
           , OdOrdQty 
           , OdExpQty 
           , OdPrdEnt 
           , OdPrdQty 
           , OdFgQty 
           , OdCalcPrc 
           , OdSalPrc 
           , OdCstPrc 
           , OdDmPrdInst 
           , OdCmPrdInst 
           , OdCmStmpInst 
           , OdSzInst 
           , OdPrtCd 
           , OdHld 
           , OdHldDesc 
           , OdVaCtg 
           , OdKt 
           , OdMulby 
           , OdFixPrc 
           , OdGldAs 
           , OdIWtEqOrd 
           , OdIWtFrOrd 
           , OdIGldRtEqOrd 
           , OdCls 
           , OdBagPcs 
           , OdDmCol 
           , OdGldAsWt 
           , ModUsr 
           , ModDt 
           , ModTime 
           , OdExpDelDt 
           , OdSfxDesc 
           , OdCoCd 
           , OdWh 
           , OdWDiaAvlblDt 
           , OdHDiaAvlblDt 

           , OdLine 
           , OdSalRem 
           , OdOmCtCd 
           , OdMinWt 
           , OdMaxWt 
           , OdBYy 
           , OdBChr 
           , OdBNo 
           , OdGmChk 
           , OdChkTol 
           , OdGrMet 
           , OdGrDia 
           , OdGrCS 
           , OdGrFin 
           , OdCrmFixPrcYN 
           , OdSubRem 
           , OdLabAs 
           , OdLabAsWt 
           , OdOmCmCd 
           , OdOmDt 
           , OdOmIdNo 
		   ,OdDmIdNo
           , OdPrtKey 
           , OdPrtFgQty 
           , OdPrtExpQty 
           , OdPicNm 
           , OdPoNo 
           , OdILabWtFrOrd 
           , OdChgPtrOnStwYN 
           , InsDt 
           , OdMinDiaTolWt 
           , OdMaxDiaTolWt 
           , OdBaseIdKey 
           , OdRefIdKey 
           , OdMrpMulby 
           , OdMrp 
           , OdDmStkYy 
           , OdDmStkChr 
           , OdDmStkNo 
			, OdJLRmCd 
			, OdJLLn1 
			, OdJLLn2 
			, OdJLLn3 
			, OdJLLotNo 
			, OdJLQty 
			, OdJLWt 
			, OdJLExpWt 
			, OdJLVchRt 
			, OdJLRateByQW 
			, OdJLVchVal 
			, OdInvAllBags 
			,OdCell
			,OdRepGrWt
			,OdDmStkCoCd
			,OdMrpDisc
			,OdFndAvlblDt 
			,OdDmTcTyp
			,OdSubCust 
			,OdMasterPrc
			,OdTag
			,OdProfit
			,OdProfitPer 
           )
		SELECT     
		   @wTargetOdTc,
           @wTargetOdYy,
           @wTargetOdChr,
           @wTargetOdNo,
		   @wOdSr,
           @wOdDmCd,
           @wOdSfx, 
           @wOdDmSz, 
		   @wOdPrdSeq,
           @wOdDeldt, 
           @wOdOrdEnt, 
           @wOdOrdQty, 
		   0 as OdExpQty,		   
		   @wOdPrdEnt,
		   @wOdPrdQty, 
		   0 as OdFgQty,
           @wOdCalcPrc,
           @wOdSalPrc,
			@wOdCstPrc,
           @wOdDmPrdInst,
           @wOdCmPrdInst,
           @wOdCmStmpInst,
		   @wOdSzInst,
           @wOdPrtCd,
           @wOdHld,
           '' as OdHldDesc,
           @wOdVaCtg
           ,@wOdKt
           ,@wOdMulBy 
           ,@wOdFixPrc
		   ,@wOdGldAs
           ,@wOdIWtEqOrd
           ,@wOdIWtFrOrd
           ,@wOdIGldRtEqOrd
		   ,@wOdCls
           ,@wOdBagPcs
           ,@wOdDmCol
           ,@wOdGldAsWt
          ,@wOdModUsr
			,convert(varchar(8),getdate(),3)
			,datepart(hh,getdate())+(convert(float,datepart(mi,getdate()))/100)
           ,@wOdExpDelDt
		   ,@wOdSfxDesc
		   ,@wTargetOdCoCd
           ,@wOdWh
           ,@wOdWDiaAvlblDt 
           ,@wOdHDiaAvlblDt 
           ,@wOdLine
           ,@wOdSalRem
           ,@wOdOmCtCd
           ,@wOdMinWt
           ,@wOdMaxWt
           ,@wOdBYy
           ,@wOdBChr
           ,@wOdBNo
           ,@wOdGmChk
           ,@wOdChkTol
           ,@wOdGrMet
           ,@wOdGrDia
           ,@wOdGrCS
           ,@wOdGrFin
           ,@wOdCrmFixPrcYN
           ,@wOdSubRem
           ,@wOdLabAs
           ,@wOdLabAsWt
           ,@wOdOmCmCd
           ,@wOdOmDt
           ,@wOdOmIdNo
           ,@wOdDmIdNo
           ,'C'
           ,0 as OdPrtFgQty
           ,0 as OdPrtExpQty
           ,@wOdPicNm
           ,@wOdPoNo
           ,@wOdILabWtFrOrd
		   ,@wOdChgPtrOnStwYN
           ,@wOdInsDt/*convert(varchar(8),getdate(),3) as InsDt*/
           ,@wOdMinDiaTolWt
           ,@wOdMaxDiaTolWt
		   ,@wOdBaseIdKey
		   ,@wOdRefIdKey 
           ,@wOdMrpMulby
           ,@wOdMrp
           ,@wOdDmStkYy 
           ,@wOdDmStkChr
           ,@wOdDmStkNo
		    ,'' as OdJLRmCd
      ,0 as OdJLLn1
      ,0 as OdJLLn2
      ,0 as OdJLLn3
      ,'' as OdJLLotNo
      ,0 as OdJLQty
      ,0.0 as OdJLWt
      ,0.0 as OdJLExpWt
      ,0.0 as OdJLVchRt
      ,'' OdJLRateByQW
      ,0.0 as OdJLVchVal
      ,@wOdInvAllBags
	  ,@wOdCell
	  ,@wOdRepGrWt
	  ,@wOdDmStkCoCd 
	  ,@wOdMrpDisc
      ,@wOdFndAvlblDt 
	  ,@wOdDmTcTyp 
	  ,@wOdSubCust 
	  ,@WOdMasterPrc 
	  ,@wOdTag
	  ,@wOdProfit
	  ,@wOdProfitPer 
	  end
	  else
	  begin
		UPDATE OrdDsg set 
           OdDmcd=@wOdDmCd,
           OdSfx=@wOdSfx, 
           OdDmSz=@wOdDmSz, 
		   OdPrdSeq=@wOdPrdSeq,
           OdDeldt=@wOdDelDt, 
           OdOrdEnt=@wOdOrdEnt,
           OdOrdQty=@wodordqty,
		   OdExpQty=0,		   
		   OdPrdEnt=@wOdPrdEnt,
		   OdPrdQty=@wOdPrdEnt, 
		   OdFgQty=0,
           OdCalcPrc=@wOdCalcPrc,
           OdSalPrc=@wOdSalPrc,
			OdCstPrc=@wOdCstPrc,
           OdDmPrdInst=@wOdDmPrdInst,
           OdCmPrdInst=@wOdCmPrdInst,
           OdCmStmpInst=@wOdCmStmpInst,
		   OdSzInst=@wOdSzInst,
           OdPrtCd=@wOdPrtCd,
           OdHld=@wOdHld,
           OdHldDesc='',
           OdVaCtg=@wOdVaCtg,
           OdKt=@wOdKt,
           OdMulBy=@wOdMulBy ,
           OdFixPrc=@wOdFixPrc,
		   OdGldAs=@wOdGldAs,
           OdIWtEqOrd=@wOdIWtEqOrd,
           OdIWtFrOrd=@wOdIWtFrOrd,
           OdIGldRtEqOrd=@wOdIGldRtEqOrd,
		   OdCls=@wOdCls,
           OdBagPcs=@wOdBagPcs,
           OdDmCol=@wOdDmCol,
           OdGldAsWt=@wOdGldAsWt,
          ModUsr=@wOdModUsr,
           OdExpDelDt=@wOdExpDelDt ,
		   OdSfxDesc=@wOdSfxDesc,
		
           OdWh=@wOdWh,
           OdWDiaAvlblDt =@wOdWDiaAvlblDt ,
           OdHDiaAvlblDt =@wOdHDiaAvlblDt ,
           OdLine=@wOdLine,
           OdSalRem=@wOdSalRem,
           OdOmCtCd=@wOdOmCtCd,
           OdMinWt=@wOdMinWt,
           OdMaxWt=@wOdMaxWt,
           OdBYy=@wOdBYy,
           OdBChr=@wOdBChr,
           OdBNo=@wOdBNo,
           odGmChk=@wOdGmChk,
           OdChkTol=@wOdChkTol,
           OdGrMet=@wOdGrMet,
           OdGrDia=@wOdGrDia,
           OdGrCS=@wOdGrCS,
           OdGrFin=@wOdGrFin,
           OdCrmFixPrcYN=@wOdCrmFixPrcYN,
          OdSubRem=@wOdSubRem,
           OdLabAs=@wOdLabAs,
           OdLabAsWt=@wOdLabAsWt,
           OdOmCmCd=@wOdOmCmCd,
           OdOmDt=@wOdOmDt,
           OdPicNm=@wOdPicNm,
           OdPoNo=@wOdPoNo,
           OdILabWtFrOrd=@wOdILabWtFrOrd,
		   OdChgPtrOnStwYN=@wOdChgPtrOnStwYN,
      
           OdMinDiaTolWt=@wOdMinDiaTolWt,
           OdMaxDiaTolWt=@wOdMaxDiaTolWt,
		   OdBaseIdKey=@wOdBaseIdKey,
		   OdRefIdKey =@wOdRefIdKey ,
           OdMrpMulby=@wOdMrpMulby,
          OdMrp=@wOdMrp,
           OdDmStkYy=@wOdDmStkYy ,
           OdDmStkChr=@wOdDmStkChr,
           OdDmStkNo=@wOdDmStkNo,
      OdInvAllBags=@wOdInvAllBags,
	  OdCell=@wOdCell,
		OdRepGrWt = @wOdRepGrWt,
	  OdDmStkCoCd =@wOdDmStkCoCd ,
	  OdMrpDisc=@wOdMrpDisc,
      OdFndAvlblDt=@wOdFndAvlblDt ,
	  OdDmTcTyp =@wOdDmTcTyp ,
	  OdSubCust =@wOdSubCust ,
	  OdMasterPrc=@wOdMasterPrc  ,
	  OdTag=@wOdTag ,
	  OdProfit=@wOdProfit ,
	  OdProfitPer=@wOdProfitPer 
		where
		          OdCoCd=@wTargetOdCoCd and OdTc=@wTargetOdTc and  OdYy =@wTargetOdYy and OdChr =@wTargetOdChr and OdNo=@wTargetOdNo and OdSr=@wOdSr 

	  end

/*	return select @@IDENTITY */

end
GOGO
ALTER TRIGGER [dbo].[OrdDsg_INUPDL] ON [dbo].[OrdDsg]     
   FOR INSERT,UPDATE,DELETE     NOT FOR REPLICATION
As    
BEGIN    
 DECLARE @ROW INT    
 DECLARE @mode char(1)    
 DECLARE @errmsg varchar(255)    
 DECLARE @modUsr char(5)
 Declare @wSelf VarChar(8)    
 Declare @wSysDb Varchar(3)    
 DECLARE @ReplAgt BIT    
 DECLARE @wCurrPrtn VarChar(1)    
 Declare @wInFob Float  
 Declare @wInFobRs Float  
 Declare @wInLclOctroiRs Float  
 Declare @wInLclVatRs Float  
 Declare @wInLclCstRs Float  
 DECLARE @wTcTyp varchar(3)   
 declare @wIsJLTc char(1)
 IF @@ROWCOUNT=0   RETURN    
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'OrdDsg' )    
  RETURN    
    
 Select @mode= case (select count('x') from inserted) when 0 then 'D'    
        else    
    case (select count('x') from deleted) when 0 then 'A' else 'C' end    
        end     
    
 Select @wSysDb=SysDb From SysDb where SysId='CRM'     
 Select @wCurrPrtn = 'C'       
 Set @wSysDb=IsNull(@wSysDb,'')     
 declare @wModUsr varchar(5)
 select @wModUsr=''
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))    
 If @ReplAgt = 1 Return
 Select @wSelf='ZSELF'     
 if @mode in ('A','C')  
 begin  
  select @wModUsr=max(inserted.ModUsr), @wTcTyp= max(PValue),@wIsJLTc = max(PValue1) from inserted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''  
  if @wTcTyp = '' select @wTcTyp = 'N'

 end  
  
 if @mode in ('D')  
 begin  
  select @wTcTyp=max(PValue), @wIsJLTc = max(PValue1) from deleted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''   
 End  
 --5717
 /******** Only One TcTyp records can be Inserted at a time.********/   
 select @Row= Count(distinct(PValue)) from inserted Join param On  
  PTyp= 'TC' and PMCd= OdTc and PScd= '' Where OdPrtKey=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'  
  GoTo ErrHndlr  
 end  
 --5717 
 /******** Only One TcTyp records can be Deleted at a time. *********/   
 select @Row= Count(distinct(PValue)) from deleted Join param on  
  PTyp= 'TC' and PMcd= OdTc and PSCd='' Where OdPrtKey=@wCurrPrtn  
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted on'  
  GoTo ErrHndlr  
 end  
   
 Select @Row= Count('x') from deleted where OdTc = 'JRI' and exists 
 (Select 'a' from Txnd where TdCoCd = OdCoCd and TdBYy = OdDmStkYy and TdBChr = OdDmStkChr and TdBNo = OdDmStkNo and (Txnd.ModDt > deleted.ModDt or (Txnd.ModDt = deleted.ModDt and Txnd.ModTime > Deleted.ModTime)))
 If @Row > 0  
 begin   
  Select @ErrMsg= 'Cannot Modify/Delete As Bag Transactions exists for Repair Bag'  
  GoTo ErrHndlr  
 end

 if @mode in ('C','D') and @wTcTyp = 'JT'
 Begin
	 Select @row= count('x') from deleted del where 
	   not exists(Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdDmCd = del.OdDmCd 
	   and ins.OdSfx= del.OdSfx and ins.OdDmSz = del.OdDmSz and ins.OdDmStkCoCd = del.OdDmStkCoCd and ins.OdDmStkYy = del.OdDmStkYy
	   and ins.OdDmStkChr = del.OdDmStkChr and ins.OdDmStkNo= del.OdDmStkNo) and
	   exists (Select 'a' from OrdRm where OrOdMstIdNo = del.OdIdNo)
	 If @row > 0     
	 begin    
	  Select @ErrMsg= 'Cannot change DmCd, Sfx, Sz or DmStkNum or delete when this is the master record for JRO transactions'  
	  goto ErrHndlr    
	 End   
 End
     
  If @Mode ='D'     
  Begin    
   Delete OrdRm From deleted, OrdRm where OdIdNo=OrOdIdNo and OdCoCd = OrCoCd and OdTc = OrTc and     
    OdYy = OrYy and OdChr = OrChr and OdNo = OrNo and OdSr= OrSr 
     
   Delete OrdLab From deleted, OrdLab where OdIdNo=OlOdIdNo and OdCoCd = OlCoCd and OdTc = OlTc and     
    OdYy = OlYy and OdChr = OlChr and OdNo = OlNo and OdSr= OlSr 
     
   Delete OrdAmd From deleted, OrdAmd where OdIdNo=OaOdIdNo and OdCoCd = OaCoCd and OdTc = OaTc and     
    OdYy = OaYy and OdChr = OaChr and OdNo = OaNo and OdSr= OaSr 
     
   Delete OrdTgt From deleted, OrdTgt where OdCoCd = OtCoCd and OdTc = OtTc and     
    OdYy = OtYy and OdChr = OtChr and OdNo = OtNo and OdSr = OtSr 
     
   Delete OrdPln From deleted, OrdPln where OdCoCd = OpCoCd and OdTc = OpTc and     
    OdYy = OpYy and OdChr = OpChr and OdNo = OpNo and OdSr = OpSr     
    
   Delete OrdCT From deleted, OrdCT where OdIdNo=OctOdIdNo and OdCoCd = OctCoCd and OdTc = OctTc and     
    OdYy = OctYy and OdChr = OctChr and OdNo = OctNo and OdSr = OctSr 
    End      
 --End    
	  
	Select @row= count('x') from inserted ins where ins.OdHld= 'Y'     
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

	  If @mode = 'C'    and (update(OdMinDiatolwt) or update(odMaxdiatolWT))
	  Begin      
	 Select @row=count('x') from inserted Where OdMinDiaTolWt > OdMaxDiaTolWt  
	 If @row > 0  
	 Begin  
	  Select @Errmsg='Minimunm Diamond Tolerance cannot be Greater Than Maximum Diamond Tolerance'  
	  Goto ErrHndlr  
	 end   
    
	 
 End    
    
  if @mode = 'A'     
  Begin    
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo 
	
    where OmReYy= '' And OmReChr= '' And OmReNo=0 And @wIsJLTc = 'N' and
    
    not exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp     
        and DmHld= 'N'  and DmCd= OdDmCd and DmSz= ''  )    
  if @row> 0     
  begin    
   select @errmsg= 'Design/Sketch Code Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from inserted     
    where @wIsJLTc = 'N' and
    not exists (Select PValue from Param where PTyp= 'DMTCTYP' and PMCd= OdDMTcTyp)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Tc Type is Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end
    
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo 
	
   where OmReYy= '' And OmReChr= '' And OmReNo=0 And  @wIsJLTc = 'N' and    
	@wIsJLTC='N' and 
    Exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp     
     and (DmHld= 'Y' ) and DmCd= OdDmCd and DmSz= OdDmSz )    
  if @row> 0     
  begin    
   select @errmsg= 'This Design/Sketch Size Is On Hold/Invalid'    
   goto ErrHndlr    
  end  
  --5717  
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) <> 'QB' And  
  (Select PMCd from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT') <> 'sri')   
  and (OdBYy <> '' or OdBChr <> '' or OdBNo <> 0)    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Should Be Blank For This Voucher Type'    
   goto ErrHndlr    
  end 

  -- Check if Not exists part with OrdDsg makes sense
  select @row= count('x') from inserted Ins where 
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('QB'))   
 and not exists (Select 'x' From Bag Where BCoCd= OdCoCd and BYy= OdBYy     
      and BChr = OdBChr And BNo= OdBNo and BLoc = 'PFG'     
      and BOdDmCd= OdDmCd )    
  /* And Not Exists (Select 'x' From OrdDsg Od Where Od.OdCoCd=ins.OdCoCd And Od.OdTc In ('QB') And     
      Od.OdBYy=Ins.OdBYy And Od.OdBChr=Ins.OdBChr And Od.OdBNo=Ins.OdBNo And Od.OdPrtKey=Ins.OdPrtKey)  */  
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Not Defined For the Design'    
   goto ErrHndlr    
  end    
end 
 if @mode in ('A','C') and @wModUsr<>'MIC'     
 Begin    
  select @row= count('x') from inserted where not exists (select 'x' from OrdMst where     
     OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and     
     OmChr= OdChr and OmNo= OdNo )    
  if @row> 0     
  begin    
   select @errmsg= 'Voucher Not Defined In Order Master'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdSr<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Design Serial Number Cannot Be <= Zero'    
   goto ErrHndlr    
  end  
  --5717
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
    
  select @row= count('x') from inserted Ins where OdDmSz<> '' and     
   Ins.OdDmSz <> IsNull((Select Del.OdDmSz From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
    Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy     
    And Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr ), '') And     
   not exists (select 'x' from Param where PTyp= 'DMSZ' and PMCd= OdDmSz And PValidYn = 'Y')    
  if @row> 0     
  begin    
   select @errmsg= 'Design Size Not Defined / Invalid.'    
   goto ErrHndlr    
  end    

  Select @row= count('x') from inserted where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and not exists (select 'x' from vParam where vPCoCd= OdCoCd and     
     vPTyp= 'PRDSEQ' and vPMCd= OdPrdSeq )    
  if @row> 0     
  begin    
   select @errmsg= 'Production Sequence Not Defined'    
   goto ErrHndlr    
  end   
    
  select @row= count('x') from inserted where OdPrdQty<  OdFgQty     
  if @row> 0     
  begin    select @errmsg= 'Production Qty Should Be >= Finished Goods Qty'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and    
   not exists (select 'x' from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd )    
  if @row> 0     
  begin    select @errmsg= 'Priority Code Not Defined'    
   goto ErrHndlr    
  end   
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPRtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdHld)    
  if @row> 0     
  begin    
   select @errmsg= 'Enter On Hold (Y/N)'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where OdPRtKey=@wCurrPRtn and OdHld= 'N' and OdHldDesc<> ''     
  if @row> 0     
  begin    
   select @errmsg= 'Holding Desc Should Be Entered Only When Design Is On Hold'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdMulBy<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Multiplying Factor Cannot Be <= Zero'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdFixPrc)    
  if @row> 0     
  begin    select @errmsg= 'Enter Fix Price (Y/N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPRtn and OdGldAs<> '' and     
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdGldAs)           
  if @row> 0     
  begin    
   select @errmsg= 'Gold As Value Not Defined'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdLabAs)    
  if @row> 0     
  begin    
   select @errmsg= 'Labour As Value Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdIWtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. Equal To Order Wt. (Y/N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdIWtFrOrd<> '' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'IWTFRORD' and PMCd= OdIWtFrOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. From Order Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdIWtFrOrd<> '' and OdIWtEqOrd= 'Y')     
  if @row> 0     
  begin    
   select @errmsg= 'Mismatch in values InvWtEqOrd and InvWtFrOrd'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'INRTORD' and PMCd= OdIGldRtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invalid value for Gold Rate Equal To Order Rate'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and     
     not exists (select 'x' from Param where PTyp='YN' and PMCd= OdCls)    
  if @row> 0     
  begin    
   select @errmsg= 'Order Closed (Y/N)'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'WH' and PMCd= OdWh)    
  if @row> 0      
  begin    
   select @errmsg= 'W/H Set Option Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdWDiaAvlblDt in('01/01/1980', '01/01/1900') or     
   OdHDiaAvlblDt in('01/01/1980', '01/01/1900') or OdFndAvlblDt in('01/01/1980', '01/01/1900'))    
  if @row> 0     
  begin    
   select @errmsg= 'Wax/ Hand Set Diamond Availability/Fnd Availability Date Cannot Be Blank'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPRtn  and OdTc <> 'JRI' and OdTc <> 'JRO' and    
    not exists (select 'x' from Param where PTyp= 'CPYRT' and PMCd= OdGmChk)    
  if @row> 0     
  begin    
   select @errmsg= 'Gram Check Option Not Defined'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
    not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChkTol)    
  if @row> 0     
  begin    
   select @errmsg= 'Should Tolerence Be Checked? (Specify Y/ N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and     
    ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and     
    ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdDmCd= del.OdDmCd)     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design'    
   goto ErrHndlr    
  end  
  --5717
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
  --5717
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
   
  --5717
  select @Row= Count('x') from inserted join InvHd on InCoCd= OdCoCd and InTc= OdTc and   
   InYy= OdYy and InChr= OdChr and InNo= OdNo where InCls= 'Y' and OdTc='FOT'  
  If @Row > 0  
  begin   
	Select @ErrMsg= 'Cannot Add/Change As Invoice Is Closed'  
	GoTo ErrHndlr  
  end  
  
    
  /******  OdVaCtg  =  PMCd ('VACTG')  ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where 
   not exists (select 'x' from Param where PTyp= 'VACTG' and PMCd= OdVaCtg)    
  if @row> 0     
  begin    
   select @errmsg= 'Value Addition Category Not Defined'    goto ErrHndlr    
  end    
    
  /******  '  =  PMCd ('KT')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and    
   not exists (select 'x' from Param where PTyp= 'KT' and PMCd= OdKt)    
  if @row> 0     
  begin    
   select @errmsg= 'Karatage Not Defined'    
   goto ErrHndlr    
  end    
    
  Select @row= count('x') from inserted Ins where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and    
   (Ins.OdDmCol= '' or     
   ( Ins.OdDmCol <> IsNull((Select Del.OdDmCol From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
      Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy And     
      Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr ), '') And     
    not exists (select 'x' from Param where PTyp= 'DMCOL' and PMCd= OdDmCol And PValidYn = 'Y')     
   ))     
  if @row> 0     
  begin     
   select @errmsg= 'Design Colour Not Defined / Invalid.'     
   goto ErrHndlr     
  End     

  Select @row= count('x') from inserted where @wIsJLTc = 'N' and 
   not exists (select 'x' from vParam where vPCoCd= OdCoCd and vPTyp= 'LINE' and vPMCd= OdLine and vPMCd<> 'ZZZ')    
  if @row> 0    begin    
   select @errmsg= 'Line Not Defined/ Line Cannot Be ''ZZZ'' '    
   goto ErrHndlr    
  end    
    
  --5717 password needed to change prdseq when bag exist
  select @row= count('x') from  deleted del Join OrdMst on OmIdNo=del.OdOmIdNo and OmCoCd= del.OdCoCd and OmTc= del.OdTc  
       and OmYy= del.OdYy and OmChr= del.OdChr and OmNo= del.OdNo where 
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdPrdSeq= del.OdPrdSeq) 
	     and (exists(Select 'x' from Bag where BOdIdNo=del.OdIdNo))    
	     AND not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change PrdSeq When Bag Exists'    
   goto ErrHndlr    
  end   
    
  select @row= count('x') from  deleted del where 
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and     
     ins.OdDmSz= del.OdDmSz and ins.OdSfx= del.OdSfx and ins.OdOrdEnt= del.OdOrdEnt and ins.OdPrdEnt= del.OdPrdEnt) 
	     and (exists(Select 'x' from Bag where BOdIdNo=del.OdIdNo))    
	     

  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design Size/ Suffix/ OrdEnt/ PrdEnt When Bag Exists'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted ins     
 join OrdMst insOm On insOm.OmIdNo=ins.OdOmIdNo       
  where 
     (Select PValue from Param where PTyp='TC' and PMCd= ins.OdTc)= 'PL' and     
     (Select count('x') from OrdDsg Od join OrdMst Om on  Om.OmIdNo=Od.OdOmIdNo where 
		Om.OmCoCd= insOm.OmCoCd and Om.OmTc= (Case When insOm.OmCmCtg = 'S' Then 'SPL' Else 'PL' End)     
    and Om.OmCmCtg= insOm.OmCmCtg and Om.OmCmCd= insOm.OmCmCd and Od.OdDmCd= ins.OdDmCd     
    and Od.OdSfx= ins.OdSfx and Od.OdDmSz= ins.OdDmSz and Om.OmCmCurCd= insOm.OmCmCurCd) > 1     
  if @row> 0     
   begin    
     select @errmsg= 'Price List Already Defined for this Cust, Curr, Dsg, Sz, Sfx combination'    
     goto ErrHndlr    
   End    

  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdDelDt= '01/01/1900' or (OdDelDt<> '01/01/1980' and     
   OdDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc     
     and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo))     
  if @row> 0     
  begin    
   select @errmsg= 'Production Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr     
  end     

  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdExpDelDt= '01/01/1900' or (OdExpDelDt<> '01/01/1980' and     
   OdExpDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd     
    and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo))     
    
  if @row> 0     
  begin    
   select @errmsg= 'Export Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr   end    
    
  Select @row= count('x') from inserted where oDtC <> 'JCN' and @wIsJLTc = 'N' and OdOrdQty < 0 or OdPrdQty< 0    
  if @row> 0     
  begin    select @errmsg= 'Order/ Prd Qty Should Be >= 0 '    
   goto ErrHndlr    
  end    
    
 select @row= count('x') from inserted where odtc <> 'JCN' and @wIsJLTc = 'N' and 
   Round(OdOrdQty, 1) < Round(OdExpQty, 1)     
  
  if @row> 0     
  begin    select @errmsg= 'Order Qty Should Be >= Export Qty '    
   goto ErrHndlr    
  end    
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
  --5717
  --select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
  -- not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and    
  --   ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdPrtKey=del.OdPrtKey and ins.OdOrdEnt= del.OdOrdEnt and ins.OdPrdEnt= del.OdPrdEnt    
  --    and ins.OdPrdSeq= del.OdPrdSeq) and     
  -- (exists (Select 'x' from Bag where BOdIdNo=OdIdNo and BCoCd= del.OdCoCd and BOdTc= del.OdTc     
  --   and BOdYy= del.OdYy and BOdChr= del.OdChr and BOdNo= del.OdNo     
  --   and BOdSr= del.OdSr and BPrtKey=del.OdPrtKey) Or OdPrtFgQty<>0 or OdPrtExpQty<>0)    
  --if @row> 0     
  --begin    
  -- select @errmsg= 'Cannot Change OrdEnt/ PrdEnt / PrdSeq When Bag Exists'    
  -- goto ErrHndlr    
  --end  
  --5717
   select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdBagPcs<= 0 and OdTc <> 'JRI' and OdTc <> 'JRO'        
  if @row> 0     
  begin    
   select @errmsg= 'Default Bag Opening Pieces Cannot Be <= Zero'    
   goto ErrHndlr    
  end 

  select @row= count('x') from inserted where OdOmCtCd<> '' and     
   not exists (select 'x' from CustMst where CmCtg= 'T' and CmCd= OdOmCtCd)    
  if @row> 0     
  begin     
   select @errmsg= 'Contractor Code Not Defined'    
   goto ErrHndlr    
  end    
    
  
  select @row= count('x') from inserted ins     
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo 
  
  where OdGrMet<>''     
   and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'MET' and PSCd= OdGrMet)     
  if @row> 0     
  begin     
   select @errmsg= 'Metal Grade not Defined'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted ins    
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo  
  where OdGrDia<>'' and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'DIA' and     
   PSCd= OdGrDia)     
  if @row> 0     
  begin     
   select @errmsg= 'Diamond Grade not Defined'    
   goto ErrHndlr    
  end    
    
  if @wTcTyp = 'JT'   
	  Begin   
	  -- rewrite 

	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmIdNo=Ins.OdOmIdNo 
	Where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
			Where Od.OdCoCd=(Case when ins.OdTc in ('JIN','LIN')  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
			
	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmIdNo=Ins.OdOmIdNo 
	Where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       ((TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
			
			Where Od.OdCoCd=Ins.OdCoCd  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo Or Od.OdDmStkNo=0)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))  
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
   
   Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo 
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
          
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
            Where Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1 =Ins.OdJLLn1 and
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.OdJLLn3    
            and (Od.OdDmStkNo=Ins.OdDmStkNo/*Or (Od.OdDmStkNo=0 Or Ins.OdDmStkNo=0)*/)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))
			
	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo 
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
	   (
       (TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
            Where Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and (Od.OdJLLotNo=Ins.OdJLLotNo or Od.OdJLLotNo = '') and Od.OdJlLn1 = Ins.OdJlLn1 and 
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))  
       )  
        )  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
   Select @Row=Count('x') from Inserted Ins Where 
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
           ), '')+',%')  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  
	Select @Row=Count('x') from Inserted Ins where
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.ODJLRmCd =Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1=Ins.OdJLLn1 
				and Od.ODJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
           ), '')+',%')  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  

	Select @row= count('x') from inserted where 
    exists(Select 'x' from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT') and Round(OdPrdQty, 1) > 0    
	   	   	   	   if @row> 0     
    begin    select @errmsg= 'Prd Qty should be 0 for Trading voucher'    
     goto ErrHndlr    
    end    

     
  end  
  

End    
--5717
if @mode = 'C'    
	  Begin      
	  select @Row= Count('x') from deleted del where     
		not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
		and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
		and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr)    
	  If @Row > 0    
	  begin     
	   Select @ErrMsg= 'Cannot Change Company Code/Tc/Yy/Chr/No/Sr'    
	   GoTo ErrHndlr    
	  end    
	 --5717
	  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and   
	  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChgPtrOnStwYN and PSCd= '')  
	 if @row> 0     
	 begin  
	  select @errmsg= 'Change Pointer On STW (Y/N)'  
	  goto ErrHndlr  
	 end   
	 --5717
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
	 --5717
	 select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and  
	  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdILabWtFrOrd and PSCd= '')  
	 if @row> 0     
	 begin  
	  select @errmsg= 'Labor Weight From Order (Y/N)'  
	  goto ErrHndlr  
	 end  
	 --5717
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
end    

 if @mode = 'D'   
 Begin    
  Select @row= count('x') from deleted where 
   (exists(Select 'x' from Bag where     
                BOdIdNo=OdIdNo ))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Export Has Taken Place / Design Under Production'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and      	 
    exists(Select 'x' from OrdLn where OdlCoCd =OdCoCd and OdlTc =OdTc and OdlYy =OdYy and OdlChr =OdChr 
    and OdlNo =OdNo and OdlSr =OdSr and OdlOdIdNo =OdIdNo )   
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Transaction Exist in OrdLn'    
   goto ErrHndlr    
  end   
  --5717
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and    
   (OdPrtExpQty<>0 Or exists(Select 'x' from Fgd where FdCoCd= OdCoCd and FdExpOdTc= OdTc and     
         FdExpOdYy = OdYy and FdExpOdChr = OdChr and FdExpOdNo = OdNo and     
         FdExpOdSr = OdSr and FdPrtKey=OdPrtKey))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Record Exists As Export Order In Fgd Table'    
   goto ErrHndlr    
  end 
    
  if @wTcTyp = 'JT'   
  Begin   
  
  /* Cannot Delete JRI if exist in Txnd */
  select @row= count('x') from deleted del where del.OdTc='JRI' and del.OdDmStkYy <> ''  and del.OdDmStkChr <> '' and del.OdDmStkNo <> ''  
    and exists(Select 'x' from Txnd tx where tx.TdCoCd =del.OdCoCd and tx.TdBYy = del.OdDmStkYy and tx.TdBChr = del.OdDmStkChr and 
		tx.TdBNo = del.OdDmStkNo and tx.ModDt >= del.ModDt and tx.ModTime >= del.ModTime)       
  if @row > 0    
  begin    
   select @errmsg= 'Cannot Delete, Bag Transaction Exists for this Repair Bag'    
   goto ErrHndlr    
  end      
  /* why is OdBaseIdKey required to be checked index is there or not */
  Select @row=Count('x') from deleted Del Where 
    exists(Select 'x' from OrdDsg Od Where   
       (Od.OdBaseIdKey=Del.OdIdNo))  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Base Voucher'  
    goto ErrHndlr  
   End  
    
   Select @row=Count('x') from deleted Del Where 
    exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and 
       Od.OdRefIdKey=Del.OdIdNo)  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Reference Voucher'  
    goto ErrHndlr  
   End  
  End   
 End    
   
 /* **** Manali Trading Module **** */  
 if @mode ='C'  
 Begin  
 if @wTcTyp = 'JT'   
  Begin  
  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where @wIsJLTc = 'N' and
   exists(Select 'x' from OrdDsg Od Where 
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
    and ins.OdDmCd=del.OdDmCd and ins.OdSfx=del.OdSfx and ins.OdDmSz=del.OdDmSz   
	and ins.OdDmStkYy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr 
    and ins.OdDmStkNo=del.OdDmStkNo
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  
      
  Select @row=Count('x') from deleted Del Where 
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and 
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'N' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and 
   ((ins.OdDmStkyy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr and ins.oddmstkno = del.oddmstkno and ins.OdDmStkNo=del.OdDmStkNo) or ins.OdDmStkNo=0 or del.OdDmStkNo=0))    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  

 End  

  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where @wIsJLTc = 'Y' and
   exists(Select 'x' from OrdDsg Od Where  (Del.OdTc='LIN' Or  Od.OdCoCd=Del.OdCoCd) and   
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
    and ins.OdJLRmCd=del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJLLn2 = del.odJLLn2 and ins.odJLLn3 = del.odJLLn3 
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  

    
  Select @row=Count('x') from deleted Del Where   
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and   
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'Y' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
     and ins.OdJLRmCd =del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJlLn2 = del.OdJlLn2 and ins.OdJLLn3 = del.ODJLLn3
     )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  
 End  

 --??End  
  
 if @mode in ('C','D')  
 begin  
 If @wTcTyp = 'FG' Or @wTcTyp = 'JT'         
 BEgin  
  update InvHd set   
  @wInFob= convert (decimal (16, 2), InFOB- IsNull((Select sum((del.OdSalPrc*del.OdOrdQty)+del.OdJLVchVal) from deleted del where del.OdCoCd= InCoCd and del.OdTc= InTc and del.OdYy= InYy   
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
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo   
  where InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD In ('FG', 'JT') and PValue= OdTc)  
         or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  
  if @wIsJLTc = 'N'
  begin
  
  update DmYyMm set  
  DDmCrQty= DDmCrQty- (select isnull(sum(Del.OdOrdQty), 0) from Deleted Del  
   Join OrdMst on   
   OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm),  
  DDmDrQty= DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkyy = DDmStkYy and del.OdDmStkChr = 
			DDmStkChr and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd    
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkCoCd =DDmStkCoCd and Deleted.OdDmStkNo= DDmStkNo and Deleted.OdDmStkYy= DDmStkYy and Deleted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
   else
   begin
  update DmYyMm set  
    DDmCrWt= DDmCrWt- (select isnull(sum(Del.OdJLWt), 0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm),  
  DDmDrWt= DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd= DJLRmCd 
   and Deleted.OdJLLotNo= DJLLotNo and Deleted.OdJLLn1= DJLLn1 and Deleted.OdJLLn2= DJLLn2 and Deleted.OdJLLn3 = DJLLn3 and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
 End  
  end
    
	if @mode in ('A','C')
	begin
		Select @row= count('x') from inserted where Inserted.OdTc='JRI' and 
			exists (Select 'a' from Bag where BCoCd=OdDmStkCoCd and BYy=OdDmStkYy and BChr = OdDmStkChr and BNo=OdDmStkNo) and
				not exists 
				(Select 'a' from Bag where 
							BCoCd = OdDmStkCoCd and BOdDmCd=OdDmCd and BOdSfx=OdSfx and BOdDmTcTyp = 'DM' and
							BOdDmSz=OdDmSz and BYy = OdDmStkYy and BChr =OdDmStkChr and BNo = OdDmStkNo and BCls='Y' and BLoc='PFG' and BRepYN='N')
 
		if @row>0 
		begin
			select @errMsg = 'Given DmStkNo not a valid Bag No For this Dsg Code-Sz-Suffix or Bag not closed or Bag not in FG or Bag already in Repair'
			goto ErrHndlr
		end
	end 
 
 If @mode in ('A','C')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'   
 Begin  

  /*** Update InvHd ***/  
  update InvHd set  
  @wInFob = convert (decimal (16, 2), InFOB+ IsNull((Select sum((ins.OdSalPrc*ins.OdOrdQty)+ins.OdJLVchVal) from inserted ins where   
   ins.OdCoCd= InCoCd and ins.OdTc= InTc and ins.OdYy= InYy   
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
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo   
  where InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD ='FG' and PValue= OdTc)  
        or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  

  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr, DDmStkNo, DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt,DdmStkCoCd)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy,OdDmStkChr,OdDmStkNo,OmFrDmLoc,convert(int,convert(varchar(6),OdOmDt,112)),0,0,@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0,OdDmStkCoCd   
   from inserted join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DCoCd= OdCoCd and DDmCd=OdDmCd   
        and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy= DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmFrDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and OdDmStkCoCd =DdmStkCoCd and 
        DYyMm= convert(int,convert(varchar(6),OdOmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
         
  /* Current Partition */   
  if @wIsJLTc ='N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd =DDmStkCoCd and Ins.OdDmStkNo = DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd =DDmStkCoCd and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdDmCd= DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkCoCd =DDmStkCoCd  
   and Inserted.OdDmStkYy = DDmStkYy and Inserted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdCoCd= DCoCd 
   and OmFrDmLoc=DLoc  and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd= DJLRmCd 
   and Inserted.OdJLLotNo= DJLLotNo and Inserted.OdJLLn1= DJLLn1 and Inserted.OdJLLn2= DJLLn2 and Inserted.OdJLLn3 = DJLLn3  and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 end  
  end
  
 if @mode in ('C','D')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'     
 Begin  
 if @wIsJLTc = 'N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst  
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst   
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd   
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkCoCd =DDmStkCoCd and Deleted.OdDmStkNo = DDmStkNo and Deleted.OdDmStkYy = DDmStkYy and 
   Deleted.OdDmStkChr = OdDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
  else
  update DmYyMm set  
  DDmCrWt=DDmCrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst  
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst   
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd = DJLRmCd and Deleted.OdJLLotNo = DJLLotNo and Deleted.OdJLLn1 = DJLLn1 and
   Deleted.OdJlLn2 = DJLLn2 and Deleted.OdJllN3 = DJLLn3  and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
 If @wTcTyp = 'JT'     
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty- IsNull((Select sum(del.OdOrdQty) from deleted del where   
     (del.OdTc='JIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt- IsNull((Select sum(del.OdJLWt) from deleted del where   
    (del.OdTc='LIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from deleted where   
    OrdDsg.OdCoCd= deleted.OdCoCd and OrdDsg.OdIdNo=deleted.OdBaseIdkey  
   and Exists (Select 'x' from TcMst Where TmCoCd=deleted.OdCoCd and TmTc=deleted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls = (case when @wIsJLTc ='N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
				else (case when OrdDsg.OdJLWt= OrdDsg.OdJLExpWt and OrdDsg.OdJLWt<> 0 then 'Y' else 'N' end)   end)
   from deleted del where OrdDsg.OdCoCd= Del.OdCoCd and OrdDsg.OdIdNo= Del.OdBaseIdKey  
      
   and Exists (Select 'x' from TcMst Where TmCoCd=del.OdCoCd and TmTc=del.OdTC and TmBaseTCs<>'')  
 
  If @mode in ('D','C')
  Begin  
   Update Bag set
		BRepYN = 'N',   
		BRepQty = 0,
		BRepINDt                ='01/01/1980',
		BRepLoc  = '',
		BRepGrWt	= 0,
		BRepOdTc = '',
		BRepOdYy = '',
		BRepOdChr = '',
		BRepOdNo = 0,
		BRepOdSr = 0
   from Deleted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRI'

   Update Bag set
   		BRepYN = 'Y'
   from Deleted
		where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRO'

  End    

 End   
  
 End   
  
 if @mode in ('A','C')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'    
 Begin  
  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr,DDmStkNo,DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt,DdmStkCoCd)
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy, OdDmStkChr, OdDmStkNo,OmToDmLoc,convert(int,convert(varchar(6),OdOmDt,112)),0,0,
					@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0,OdDmStkCoCd  
     
   from inserted join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DCoCd= OdCoCd and DDmCd=OdDmCd   
       and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy = DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmToDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and OdDmStkCoCd =DdmStkCoCd and 
       DYyMm= convert(int,convert(varchar(6),OdOmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  
  if @wIsJLTc = 'N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd= DDmStkCoCd and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd= DDmStkCoCd and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdDmCd=DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkCoCd= DDmStkCoCd and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkChr = DDmStkChr and Inserted.OdDmStkYy = DDmStkYy 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd = DJLRmCd and Inserted.ODJLLotNo = DJLLotNo and Inserted.ODJLLn1 = DJLLn1 and Inserted.ODJLLn2 = DJLLn2 and 
   Inserted.ODJLLn3 = DJLLn3 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 If @wTcTyp = 'JT'   
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty + IsNull((Select sum(Ins.OdOrdQty) from Inserted Ins where   
   (ins.OdTc='JIN' Or Ins.OdCoCd= OrdDsg.OdCoCd)  and Ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt+ IsNull((Select sum(ins.OdJLWt) from Inserted Ins where   
   (ins.OdTc='LIN' Or ins.OdCoCd= OrdDsg.OdCoCd) and ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from Inserted where 
    OrdDsg.OdCoCd= Inserted.OdCoCd and OrdDsg.OdIdNo=Inserted.OdBaseIdkey  
    and Exists (Select 'x' from TcMst Where TmCoCd=Inserted.OdCoCd and TmTc=Inserted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls =(case when @wIsJLTc = 'N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
			else (case when OrdDsg.OdJLWt = OrdDsg.ODJLExpWt and OrdDsg.ODJLWt <> 0 then 'Y' else 'N' end) end)
   from Inserted Ins where (ins.OdTc = 'LIN' or ins.OdTc='JIN' Or  OrdDsg.OdCoCd= Ins.OdCoCd)   
   and OrdDsg.OdIdNo= Ins.OdBaseIdKey and Exists (Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTC and TmBaseTCs<>'')  
  if @mode in ('A','C')  
  Begin   
   Update Bag set   
		BRepYN = 'Y',
		BRepQty = OdOrdEnt,
		BRepINDt = getdate(),
		/* (Select OmDt From OrdMst Where OmCoCd= OdCoCd And OmTc = OdTc   
		And OmYy = OdYy And OmChr= OdChr And OmNo= OdNo And OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo)*/
		BRepLoc  = (Select HDefRepBagLoc from Head where HCoCd = Inserted.OdCoCd),
		BRepGrWt = OdRepGrWt,
		BRepOdTc = OdTc,
		BRepOdYy = OdYy,
		BRepOdChr = OdChr,
		BRepOdNo = OdNo,
		BRepOdSr = OdSr
   from Inserted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRI'

   Update Bag set
		BRepYN= 'N'
   from Inserted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRO'
  End  
 End  
  
 End    
  
  
if @mode in ('C','A') and @wTcTyp = 'JT'   
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') from Inserted Ins   
   where exists(Select 'x' from OrdDsg B Where  
      B.OdIdNo=Ins.OdBaseIdkey and B.OdOrdQty < B.OdExpQty)  
    and (exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>''))  
  If @Row> 0  
  begin   
   select @errmsg= 'Order Qty Should Be >= Matched Qty for Base Voucher'    
   goto ErrHndlr    
  end
 end
else
begin
 Select @Row=Count('x') from Inserted Ins   
   where 
    exists(Select 'x' from OrdDsg B Where  
      B.OdIdNo=Ins.OdBaseIdkey and cast(B.OdJLWt as decimal(10,3))< cast(B.OdJLExpWt as decimal(10,3)))  
    and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>'')  
 If @Row> 0  
  begin   
   select @errmsg= 'Order Wt Should Be >= Matched Wt for Base Voucher'    
   goto ErrHndlr    
  end
 end   
end    
  
If @mode in ('C','A') and  (@wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then 
	  (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) 
	  else (case when OmFrDmDc = 'C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Ins.OdCoCd= DCoCd and Ins.OdDmCd= DDmCd and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and   
    Ins.OdDmStkCoCd =DdmStkCoCd and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkNo = DDmStkNo and 
     (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Desing Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
end
else
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrWt)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.OdJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3=DJLLn3 and
	(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Size'  
     GoTo ErrHndlr  
     end  

end
end  
If @mode in ('D') and  (@wTcTyp = 'FG' or @wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 -- rewrite checking JCN why and remove FG and see what to simplify 
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd = Del.OdCoCd and LocTyp =(Case PValue When 'JT' Then 'S' End)  and   
            LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then (case when OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and   
    Del.OdDmStkCoCd =DdmStkCoCd and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo and 
    (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Design Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
	end
else
begin
 -- rewrite checking JCN why and remove FG and see what to simplify 
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Del.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End)  and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrwT)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Del.OdCoCd= DCoCd and Del.ODJLRmCd = DJLRmCd and Del.ODJLLotNo = DJLLotNo and Del.ODJLLn1 = DJLLn1 and Del.ODJLLn2 = DJLLn2 and Del.ODJLLn3 = DJLLn3 and
(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Sz'  
     GoTo ErrHndlr  
     end  
end

end  
return    
ErrHndlr:    
 begin   select @errmsg='|'+@errmsg    
  raiserror(@errmsg,16,1)   rollback transaction     
 end    
end   



GOGO
ALTER TRIGGER [dbo].[OrdDsg_INUPDL] ON [dbo].[OrdDsg]     
   FOR INSERT,UPDATE,DELETE     NOT FOR REPLICATION
As    
BEGIN    
 DECLARE @ROW INT    
 DECLARE @mode char(1)    
 DECLARE @errmsg varchar(255)    
 DECLARE @modUsr char(5)
 Declare @wSelf VarChar(8)    
 Declare @wSysDb Varchar(3)    
 DECLARE @ReplAgt BIT    
 DECLARE @wCurrPrtn VarChar(1)    
 Declare @wInFob Float  
 Declare @wInFobRs Float  
 Declare @wInLclOctroiRs Float  
 Declare @wInLclVatRs Float  
 Declare @wInLclCstRs Float  
 DECLARE @wTcTyp varchar(3)   
 declare @wIsJLTc char(1)
 IF @@ROWCOUNT=0   RETURN    
 If Exists(Select 'x' From TrgDisa Where TdTblNm = 'OrdDsg' )    
  RETURN    
    
 Select @mode= case (select count('x') from inserted) when 0 then 'D'    
        else    
    case (select count('x') from deleted) when 0 then 'A' else 'C' end    
        end     
    
 Select @wSysDb=SysDb From SysDb where SysId='CRM'     
 Select @wCurrPrtn = 'C'       
 Set @wSysDb=IsNull(@wSysDb,'')     
 declare @wModUsr varchar(5)
 select @wModUsr=''
 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))    
 If @ReplAgt = 1 Return
 Select @wSelf='ZSELF'     
 if @mode in ('A','C')  
 begin  
  select @wModUsr=max(inserted.ModUsr), @wTcTyp= max(PValue),@wIsJLTc = max(PValue1) from inserted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''  
  if @wTcTyp = '' select @wTcTyp = 'N'

 end  
  
 if @mode in ('D')  
 begin  
  select @wTcTyp=max(PValue), @wIsJLTc = max(PValue1) from deleted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''   
 End  
 --5717
 /******** Only One TcTyp records can be Inserted at a time.********/   
 select @Row= Count(distinct(PValue)) from inserted Join param On  
  PTyp= 'TC' and PMCd= OdTc and PScd= '' Where OdPrtKey=@wCurrPrtn   
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted On'  
  GoTo ErrHndlr  
 end  
 --5717 
 /******** Only One TcTyp records can be Deleted at a time. *********/   
 select @Row= Count(distinct(PValue)) from deleted Join param on  
  PTyp= 'TC' and PMcd= OdTc and PSCd='' Where OdPrtKey=@wCurrPrtn  
 If @Row> 1  
 begin  
  Select @ErrMsg= 'Transactions of only One TcType can be Acted on'  
  GoTo ErrHndlr  
 end  
   
 Select @Row= Count('x') from deleted where OdTc = 'JRI' and exists 
 (Select 'a' from Txnd where TdCoCd = OdCoCd and TdBYy = OdDmStkYy and TdBChr = OdDmStkChr and TdBNo = OdDmStkNo and (Txnd.ModDt > deleted.ModDt or (Txnd.ModDt = deleted.ModDt and Txnd.ModTime > Deleted.ModTime)))
 If @Row > 0  
 begin   
  Select @ErrMsg= 'Cannot Modify/Delete As Bag Transactions exists for Repair Bag'  
  GoTo ErrHndlr  
 end

 if @mode in ('C','D') and @wTcTyp = 'JT'
 Begin
	 Select @row= count('x') from deleted del where 
	   not exists(Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdDmCd = del.OdDmCd 
	   and ins.OdSfx= del.OdSfx and ins.OdDmSz = del.OdDmSz and ins.OdDmStkCoCd = del.OdDmStkCoCd and ins.OdDmStkYy = del.OdDmStkYy
	   and ins.OdDmStkChr = del.OdDmStkChr and ins.OdDmStkNo= del.OdDmStkNo) and
	   exists (Select 'a' from OrdRm where OrOdMstIdNo = del.OdIdNo)
	 If @row > 0     
	 begin    
	  Select @ErrMsg= 'Cannot change DmCd, Sfx, Sz or DmStkNum or delete when this is the master record for JRO transactions'  
	  goto ErrHndlr    
	 End   
 End
     
  If @Mode ='D'     
  Begin    
   Delete OrdRm From deleted, OrdRm where OdIdNo=OrOdIdNo and OdCoCd = OrCoCd and OdTc = OrTc and     
    OdYy = OrYy and OdChr = OrChr and OdNo = OrNo and OdSr= OrSr 
     
   Delete OrdLab From deleted, OrdLab where OdIdNo=OlOdIdNo and OdCoCd = OlCoCd and OdTc = OlTc and     
    OdYy = OlYy and OdChr = OlChr and OdNo = OlNo and OdSr= OlSr 
     
   Delete OrdAmd From deleted, OrdAmd where OdIdNo=OaOdIdNo and OdCoCd = OaCoCd and OdTc = OaTc and     
    OdYy = OaYy and OdChr = OaChr and OdNo = OaNo and OdSr= OaSr 
     
   Delete OrdTgt From deleted, OrdTgt where OdCoCd = OtCoCd and OdTc = OtTc and     
    OdYy = OtYy and OdChr = OtChr and OdNo = OtNo and OdSr = OtSr 
     
   Delete OrdPln From deleted, OrdPln where OdCoCd = OpCoCd and OdTc = OpTc and     
    OdYy = OpYy and OdChr = OpChr and OdNo = OpNo and OdSr = OpSr     
    
   Delete OrdCT From deleted, OrdCT where OdIdNo=OctOdIdNo and OdCoCd = OctCoCd and OdTc = OctTc and     
    OdYy = OctYy and OdChr = OctChr and OdNo = OctNo and OdSr = OctSr 
    End      
 --End    
	  
	Select @row= count('x') from inserted ins where ins.OdHld= 'Y'     
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

	  If @mode = 'C'    and (update(OdMinDiatolwt) or update(odMaxdiatolWT))
	  Begin      
	 Select @row=count('x') from inserted Where OdMinDiaTolWt > OdMaxDiaTolWt  
	 If @row > 0  
	 Begin  
	  Select @Errmsg='Minimunm Diamond Tolerance cannot be Greater Than Maximum Diamond Tolerance'  
	  Goto ErrHndlr  
	 end   
    
	 
 End    
    
  if @mode = 'A'     
  Begin    
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo 
	
    where OmReYy= '' And OmReChr= '' And OmReNo=0 And @wIsJLTc = 'N' and
    
    not exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp     
        and DmHld= 'N'  and DmCd= OdDmCd and DmSz= ''  )    
  if @row> 0     
  begin    
   select @errmsg= 'Design/Sketch Code Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from inserted     
    where @wIsJLTc = 'N' and
    not exists (Select PValue from Param where PTyp= 'DMTCTYP' and PMCd= OdDMTcTyp)    
  if @row> 0     
  begin    
   select @errmsg= 'Design Tc Type is Not Defined Or Is On Hold'    
   goto ErrHndlr    
  end
    
  select @row= count('x') from inserted     
    Join OrdMst On OmIdNo=OdOmIdNo and OmCoCd=OdCoCd And OmTc=OdTc And OmYy=OdYy And OmChr=OdChr And OmNo=OdNo 
	
   where OmReYy= '' And OmReChr= '' And OmReNo=0 And  @wIsJLTc = 'N' and    
	@wIsJLTC='N' and 
    Exists (select 'x' from DsgMst where DmIdNo=OdDmIdNo and DmTcTyp= OdDmTcTyp     
     and (DmHld= 'Y' ) and DmCd= OdDmCd and DmSz= OdDmSz )    
  if @row> 0     
  begin    
   select @errmsg= 'This Design/Sketch Size Is On Hold/Invalid'    
   goto ErrHndlr    
  end  
  --5717  
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and    
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) <> 'QB' And  
  (Select PMCd from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT') <> 'sri')   
  and (OdBYy <> '' or OdBChr <> '' or OdBNo <> 0)    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Should Be Blank For This Voucher Type'    
   goto ErrHndlr    
  end 

  -- Check if Not exists part with OrdDsg makes sense
  select @row= count('x') from inserted Ins where 
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('QB'))   
 and not exists (Select 'x' From Bag Where BCoCd= OdCoCd and BYy= OdBYy     
      and BChr = OdBChr And BNo= OdBNo and BLoc = 'PFG'     
      and BOdDmCd= OdDmCd )    
  /* And Not Exists (Select 'x' From OrdDsg Od Where Od.OdCoCd=ins.OdCoCd And Od.OdTc In ('QB') And     
      Od.OdBYy=Ins.OdBYy And Od.OdBChr=Ins.OdBChr And Od.OdBNo=Ins.OdBNo And Od.OdPrtKey=Ins.OdPrtKey)  */  
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Not Defined For the Design'    
   goto ErrHndlr    
  end    
end 
 if @mode in ('A','C') and @wModUsr<>'MIC'     
 Begin    
  select @row= count('x') from inserted where not exists (select 'x' from OrdMst where     
     OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc and OmYy= OdYy and     
     OmChr= OdChr and OmNo= OdNo )    
  if @row> 0     
  begin    
   select @errmsg= 'Voucher Not Defined In Order Master'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and OdSr<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Design Serial Number Cannot Be <= Zero'    
   goto ErrHndlr    
  end  
  --5717
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
    
  select @row= count('x') from inserted Ins where OdDmSz<> '' and     
   Ins.OdDmSz <> IsNull((Select Del.OdDmSz From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
    Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy     
    And Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr ), '') And     
   not exists (select 'x' from Param where PTyp= 'DMSZ' and PMCd= OdDmSz And PValidYn = 'Y')    
  if @row> 0     
  begin    
   select @errmsg= 'Design Size Not Defined / Invalid.'    
   goto ErrHndlr    
  end    

  Select @row= count('x') from inserted where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and not exists (select 'x' from vParam where vPCoCd= OdCoCd and     
     vPTyp= 'PRDSEQ' and vPMCd= OdPrdSeq )    
  if @row> 0     
  begin    
   select @errmsg= 'Production Sequence Not Defined'    
   goto ErrHndlr    
  end   
    
  select @row= count('x') from inserted where OdPrdQty<  OdFgQty     
  if @row> 0     
  begin    select @errmsg= 'Production Qty Should Be >= Finished Goods Qty'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and    
   not exists (select 'x' from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd )    
  if @row> 0     
  begin    select @errmsg= 'Priority Code Not Defined'    
   goto ErrHndlr    
  end   
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPRtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdHld)    
  if @row> 0     
  begin    
   select @errmsg= 'Enter On Hold (Y/N)'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where OdPRtKey=@wCurrPRtn and OdHld= 'N' and OdHldDesc<> ''     
  if @row> 0     
  begin    
   select @errmsg= 'Holding Desc Should Be Entered Only When Design Is On Hold'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdMulBy<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Multiplying Factor Cannot Be <= Zero'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdFixPrc)    
  if @row> 0     
  begin    select @errmsg= 'Enter Fix Price (Y/N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPRtn and OdGldAs<> '' and     
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdGldAs)           
  if @row> 0     
  begin    
   select @errmsg= 'Gold As Value Not Defined'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdLabAs)    
  if @row> 0     
  begin    
   select @errmsg= 'Labour As Value Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdIWtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. Equal To Order Wt. (Y/N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdIWtFrOrd<> '' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'IWTFRORD' and PMCd= OdIWtFrOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. From Order Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdIWtFrOrd<> '' and OdIWtEqOrd= 'Y')     
  if @row> 0     
  begin    
   select @errmsg= 'Mismatch in values InvWtEqOrd and InvWtFrOrd'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'INRTORD' and PMCd= OdIGldRtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invalid value for Gold Rate Equal To Order Rate'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and     
     not exists (select 'x' from Param where PTyp='YN' and PMCd= OdCls)    
  if @row> 0     
  begin    
   select @errmsg= 'Order Closed (Y/N)'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'WH' and PMCd= OdWh)    
  if @row> 0      
  begin    
   select @errmsg= 'W/H Set Option Not Defined'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdWDiaAvlblDt in('01/01/1980', '01/01/1900') or     
   OdHDiaAvlblDt in('01/01/1980', '01/01/1900') or OdFndAvlblDt in('01/01/1980', '01/01/1900'))    
  if @row> 0     
  begin    
   select @errmsg= 'Wax/ Hand Set Diamond Availability/Fnd Availability Date Cannot Be Blank'    
   goto ErrHndlr    
  end  
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPRtn  and OdTc <> 'JRI' and OdTc <> 'JRO' and    
    not exists (select 'x' from Param where PTyp= 'CPYRT' and PMCd= OdGmChk)    
  if @row> 0     
  begin    
   select @errmsg= 'Gram Check Option Not Defined'    
   goto ErrHndlr    
  end
  --5717
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
    not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChkTol)    
  if @row> 0     
  begin    
   select @errmsg= 'Should Tolerence Be Checked? (Specify Y/ N)'    
   goto ErrHndlr    
  end 
  --5717
  select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and     
    ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and     
    ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdDmCd= del.OdDmCd)     
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design'    
   goto ErrHndlr    
  end  
  --5717
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
  --5717
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
   
  --5717
  select @Row= Count('x') from inserted join InvHd on InCoCd= OdCoCd and InTc= OdTc and   
   InYy= OdYy and InChr= OdChr and InNo= OdNo where InCls= 'Y' and OdTc='FOT'  
  If @Row > 0  
  begin   
	Select @ErrMsg= 'Cannot Add/Change As Invoice Is Closed'  
	GoTo ErrHndlr  
  end  
  
    
  /******  OdVaCtg  =  PMCd ('VACTG')  ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where 
   not exists (select 'x' from Param where PTyp= 'VACTG' and PMCd= OdVaCtg)    
  if @row> 0     
  begin    
   select @errmsg= 'Value Addition Category Not Defined'    goto ErrHndlr    
  end    
    
  /******  '  =  PMCd ('KT')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and    
   not exists (select 'x' from Param where PTyp= 'KT' and PMCd= OdKt)    
  if @row> 0     
  begin    
   select @errmsg= 'Karatage Not Defined'    
   goto ErrHndlr    
  end    
    
  Select @row= count('x') from inserted Ins where @wIsJLTc = 'N' and OdTc <> 'JRI' and OdTc <> 'JRO' and    
   (Ins.OdDmCol= '' or     
   ( Ins.OdDmCol <> IsNull((Select Del.OdDmCol From Deleted Del Where Ins.OdIdNo=Del.OdIdNo and     
      Ins.OdCoCd = Del.OdCoCd And Ins.OdTc = Del.OdTc And Ins.OdYy = Del.OdYy And     
      Ins.OdChr = Del.OdChr And Ins.OdNo = Del.OdNo And Ins.OdSr = Del.OdSr ), '') And     
    not exists (select 'x' from Param where PTyp= 'DMCOL' and PMCd= OdDmCol And PValidYn = 'Y')     
   ))     
  if @row> 0     
  begin     
   select @errmsg= 'Design Colour Not Defined / Invalid.'     
   goto ErrHndlr     
  End     

  Select @row= count('x') from inserted where @wIsJLTc = 'N' and 
   not exists (select 'x' from vParam where vPCoCd= OdCoCd and vPTyp= 'LINE' and vPMCd= OdLine and vPMCd<> 'ZZZ')    
  if @row> 0    begin    
   select @errmsg= 'Line Not Defined/ Line Cannot Be ''ZZZ'' '    
   goto ErrHndlr    
  end    
    
  --5717 password needed to change prdseq when bag exist
  select @row= count('x') from  deleted del Join OrdMst on OmIdNo=del.OdOmIdNo and OmCoCd= del.OdCoCd and OmTc= del.OdTc  
       and OmYy= del.OdYy and OmChr= del.OdChr and OmNo= del.OdNo where 
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdPrdSeq= del.OdPrdSeq) 
	     and (exists(Select 'x' from Bag where BOdIdNo=del.OdIdNo))    
	     AND not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change PrdSeq When Bag Exists'    
   goto ErrHndlr    
  end   
    
  select @row= count('x') from  deleted del where 
   not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and     
     ins.OdDmSz= del.OdDmSz and ins.OdSfx= del.OdSfx and ins.OdOrdEnt= del.OdOrdEnt and ins.OdPrdEnt= del.OdPrdEnt) 
	     and (exists(Select 'x' from Bag where BOdIdNo=del.OdIdNo))    
	     

  if @row> 0     
  begin    
   select @errmsg= 'Cannot Change Design Size/ Suffix/ OrdEnt/ PrdEnt When Bag Exists'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted ins     
 join OrdMst insOm On insOm.OmIdNo=ins.OdOmIdNo       
  where 
     (Select PValue from Param where PTyp='TC' and PMCd= ins.OdTc)= 'PL' and     
     (Select count('x') from OrdDsg Od join OrdMst Om on  Om.OmIdNo=Od.OdOmIdNo where 
		Om.OmCoCd= insOm.OmCoCd and Om.OmTc= (Case When insOm.OmCmCtg = 'S' Then 'SPL' Else 'PL' End)     
    and Om.OmCmCtg= insOm.OmCmCtg and Om.OmCmCd= insOm.OmCmCd and Od.OdDmCd= ins.OdDmCd     
    and Od.OdSfx= ins.OdSfx and Od.OdDmSz= ins.OdDmSz and Om.OmCmCurCd= insOm.OmCmCurCd) > 1     
  if @row> 0     
   begin    
     select @errmsg= 'Price List Already Defined for this Cust, Curr, Dsg, Sz, Sfx combination'    
     goto ErrHndlr    
   End    

  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdDelDt= '01/01/1900' or (OdDelDt<> '01/01/1980' and     
   OdDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd and OmTc= OdTc     
     and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo))     
  if @row> 0     
  begin    
   select @errmsg= 'Production Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr     
  end     

  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdExpDelDt= '01/01/1900' or (OdExpDelDt<> '01/01/1980' and     
   OdExpDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd     
    and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo))     
    
  if @row> 0     
  begin    
   select @errmsg= 'Export Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr   end    
    
  Select @row= count('x') from inserted where oDtC <> 'JCN' and @wIsJLTc = 'N' and OdOrdQty < 0 or OdPrdQty< 0    
  if @row> 0     
  begin    select @errmsg= 'Order/ Prd Qty Should Be >= 0 '    
   goto ErrHndlr    
  end    
    
 select @row= count('x') from inserted where odtc <> 'JCN' and @wIsJLTc = 'N' and 
   Round(OdOrdQty, 1) < Round(OdExpQty, 1)     
  
  if @row> 0     
  begin    select @errmsg= 'Order Qty Should Be >= Export Qty '    
   goto ErrHndlr    
  end    
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
  --5717
  --select @row= count('x') from  deleted del where del.OdPRtKey=@wCurrPrtn and    
  -- not exists (select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr and    
  --   ins.OdNo= del.OdNo and ins.OdSr= del.OdSr and ins.OdPrtKey=del.OdPrtKey and ins.OdOrdEnt= del.OdOrdEnt and ins.OdPrdEnt= del.OdPrdEnt    
  --    and ins.OdPrdSeq= del.OdPrdSeq) and     
  -- (exists (Select 'x' from Bag where BOdIdNo=OdIdNo and BCoCd= del.OdCoCd and BOdTc= del.OdTc     
  --   and BOdYy= del.OdYy and BOdChr= del.OdChr and BOdNo= del.OdNo     
  --   and BOdSr= del.OdSr and BPrtKey=del.OdPrtKey) Or OdPrtFgQty<>0 or OdPrtExpQty<>0)    
  --if @row> 0     
  --begin    
  -- select @errmsg= 'Cannot Change OrdEnt/ PrdEnt / PrdSeq When Bag Exists'    
  -- goto ErrHndlr    
  --end  
  --5717
   select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdBagPcs<= 0 and OdTc <> 'JRI' and OdTc <> 'JRO'        
  if @row> 0     
  begin    
   select @errmsg= 'Default Bag Opening Pieces Cannot Be <= Zero'    
   goto ErrHndlr    
  end 

  select @row= count('x') from inserted where OdOmCtCd<> '' and     
   not exists (select 'x' from CustMst where CmCtg= 'T' and CmCd= OdOmCtCd)    
  if @row> 0     
  begin     
   select @errmsg= 'Contractor Code Not Defined'    
   goto ErrHndlr    
  end    
    
  
  select @row= count('x') from inserted ins     
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo 
  
  where OdGrMet<>''     
   and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'MET' and PSCd= OdGrMet)     
  if @row> 0     
  begin     
   select @errmsg= 'Metal Grade not Defined'    
   goto ErrHndlr    
  end    
    
  select @row= count('x') from inserted ins    
  Join OrdMst On OdOmIdNo=OmIdNo and OdCoCd=OmCoCd and OdTc=OmTc and OdYy=OmYy and OdChr=OmChr and OdNo=OmNo  
  where OdGrDia<>'' and Not Exists(Select 'x' From Param where PTyp= 'GRDCD' and PMCd= 'DIA' and     
   PSCd= OdGrDia)     
  if @row> 0     
  begin     
   select @errmsg= 'Diamond Grade not Defined'    
   goto ErrHndlr    
  end    
    
  if @wTcTyp = 'JT'   
	  Begin   
	  -- rewrite 

	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmIdNo=Ins.OdOmIdNo 
	Where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
			Where Od.OdCoCd=(Case when ins.OdTc in ('JIN','LIN')  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
			
	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmIdNo=Ins.OdOmIdNo 
	Where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       ((TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
			
			Where Od.OdCoCd=Ins.OdCoCd  
            and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
            and (Od.OdDmStkNo=Ins.OdDmStkNo Or Od.OdDmStkNo=0)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))  
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
   
   Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo 
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
          
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
            Where Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1 =Ins.OdJLLn1 and
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.OdJLLn3    
            and (Od.OdDmStkNo=Ins.OdDmStkNo/*Or (Od.OdDmStkNo=0 Or Ins.OdDmStkNo=0)*/)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))))
			
	Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo 
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
	   (
       (TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo 
            Where Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and (Od.OdJLLotNo=Ins.OdJLLotNo or Od.OdJLLotNo = '') and Od.OdJlLn1 = Ins.OdJlLn1 and 
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), ''))  
       )  
        )  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Base Voucher'  
    GoTo ErrHndlr  
   end  
   Select @Row=Count('x') from Inserted Ins Where 
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.OdDmCd=Ins.OdDmCd and Od.OdSfx=Ins.OdSfx and Od.OdDmSz=Ins.OdDmSz   
           ), '')+',%')  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  
	Select @Row=Count('x') from Inserted Ins where
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.ODJLRmCd =Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1=Ins.OdJLLn1 
				and Od.ODJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
           ), '')+',%')  
	   	   	   	   	   If @Row > 0  
   Begin  
    Select @ErrMsg= 'Invalid Reference Voucher'  
    GoTo ErrHndlr  
   end  

	Select @row= count('x') from inserted where 
    exists(Select 'x' from Param Where PTyp='TC' and PMCd=OdTc and PValue='JT') and Round(OdPrdQty, 1) > 0    
	   	   	   	   if @row> 0     
    begin    select @errmsg= 'Prd Qty should be 0 for Trading voucher'    
     goto ErrHndlr    
    end    

     
  end  
  

End    
--5717
if @mode = 'C'    
	  Begin      
	  select @Row= Count('x') from deleted del where     
		not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
		and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
		and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr)    
	  If @Row > 0    
	  begin     
	   Select @ErrMsg= 'Cannot Change Company Code/Tc/Yy/Chr/No/Sr'    
	   GoTo ErrHndlr    
	  end    
	 --5717
	  select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and   
	  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdChgPtrOnStwYN and PSCd= '')  
	 if @row> 0     
	 begin  
	  select @errmsg= 'Change Pointer On STW (Y/N)'  
	  goto ErrHndlr  
	 end   
	 --5717
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
	 --5717
	 select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and  
	  not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdILabWtFrOrd and PSCd= '')  
	 if @row> 0     
	 begin  
	  select @errmsg= 'Labor Weight From Order (Y/N)'  
	  goto ErrHndlr  
	 end  
	 --5717
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
end    

 if @mode = 'D'   
 Begin    
  Select @row= count('x') from deleted where 
   (exists(Select 'x' from Bag where     
                BOdIdNo=OdIdNo ))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Export Has Taken Place / Design Under Production'    
   goto ErrHndlr    
  end    
  --5717
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and      	 
    exists(Select 'x' from OrdLn where OdlCoCd =OdCoCd and OdlTc =OdTc and OdlYy =OdYy and OdlChr =OdChr 
    and OdlNo =OdNo and OdlSr =OdSr and OdlOdIdNo =OdIdNo )   
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete Transaction Exist in OrdLn'    
   goto ErrHndlr    
  end   
  --5717
  select @row= count('x') from deleted where OdPrtKey=@wCurrPrtn and    
   (OdPrtExpQty<>0 Or exists(Select 'x' from Fgd where FdCoCd= OdCoCd and FdExpOdTc= OdTc and     
         FdExpOdYy = OdYy and FdExpOdChr = OdChr and FdExpOdNo = OdNo and     
         FdExpOdSr = OdSr and FdPrtKey=OdPrtKey))    
  if @row> 0     
  begin    
   select @errmsg= 'Cannot Delete As Record Exists As Export Order In Fgd Table'    
   goto ErrHndlr    
  end 
    
  if @wTcTyp = 'JT'   
  Begin   
  
  /* Cannot Delete JRI if exist in Txnd */
  select @row= count('x') from deleted del where del.OdTc='JRI' and del.OdDmStkYy <> ''  and del.OdDmStkChr <> '' and del.OdDmStkNo <> ''  
    and exists(Select 'x' from Txnd tx where tx.TdCoCd =del.OdCoCd and tx.TdBYy = del.OdDmStkYy and tx.TdBChr = del.OdDmStkChr and 
		tx.TdBNo = del.OdDmStkNo and tx.ModDt >= del.ModDt and tx.ModTime >= del.ModTime)       
  if @row > 0    
  begin    
   select @errmsg= 'Cannot Delete, Bag Transaction Exists for this Repair Bag'    
   goto ErrHndlr    
  end      
  /* why is OdBaseIdKey required to be checked index is there or not */
  Select @row=Count('x') from deleted Del Where 
    exists(Select 'x' from OrdDsg Od Where   
       (Od.OdBaseIdKey=Del.OdIdNo))  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Base Voucher'  
    goto ErrHndlr  
   End  
    
   Select @row=Count('x') from deleted Del Where 
    exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and 
       Od.OdRefIdKey=Del.OdIdNo)  
   If @row > 0  
   Begin  
    Select @errmsg='Cannot Delete as Order Serial Is Used as Reference Voucher'  
    goto ErrHndlr  
   End  
  End   
 End    
   
 /* **** Manali Trading Module **** */  
 if @mode ='C'  
 Begin  
 if @wTcTyp = 'JT'   
  Begin  
  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where @wIsJLTc = 'N' and
   exists(Select 'x' from OrdDsg Od Where 
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
    and ins.OdDmCd=del.OdDmCd and ins.OdSfx=del.OdSfx and ins.OdDmSz=del.OdDmSz   
	and ins.OdDmStkYy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr 
    and ins.OdDmStkNo=del.OdDmStkNo
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  
      
  Select @row=Count('x') from deleted Del Where 
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and 
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'N' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and 
   ((ins.OdDmStkyy = del.oddmstkyy and ins.oddmstkchr = del.oddmstkchr and ins.oddmstkno = del.oddmstkno and ins.OdDmStkNo=del.OdDmStkNo) or ins.OdDmStkNo=0 or del.OdDmStkNo=0))    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change Design/Size/Suffix/DmId as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  

 End  

  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where @wIsJLTc = 'Y' and
   exists(Select 'x' from OrdDsg Od Where  (Del.OdTc='LIN' Or  Od.OdCoCd=Del.OdCoCd) and   
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
    and ins.OdJLRmCd=del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJLLn2 = del.odJLLn2 and ins.odJLLn3 = del.odJLLn3 
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  

    
  Select @row=Count('x') from deleted Del Where   
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and   
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'Y' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo 
     and ins.OdJLRmCd =del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJlLn2 = del.OdJlLn2 and ins.OdJLLn3 = del.ODJLLn3
     )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Reference Voucher'  
   goto ErrHndlr  
  End  
 End  

 --??End  
  
 if @mode in ('C','D')  
 begin  
 If @wTcTyp = 'FG' Or @wTcTyp = 'JT'         
 BEgin  
  update InvHd set   
  @wInFob= convert (decimal (16, 2), InFOB- IsNull((Select sum((del.OdSalPrc*del.OdOrdQty)+del.OdJLVchVal) from deleted del where del.OdCoCd= InCoCd and del.OdTc= InTc and del.OdYy= InYy   
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
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo   
  where InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD In ('FG', 'JT') and PValue= OdTc)  
         or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  
  if @wIsJLTc = 'N'
  begin
  
  update DmYyMm set  
  DDmCrQty= DDmCrQty- (select isnull(sum(Del.OdOrdQty), 0) from Deleted Del  
   Join OrdMst on   
   OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm),  
  DDmDrQty= DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkyy = DDmStkYy and del.OdDmStkChr = 
			DDmStkChr and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd    
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkCoCd =DDmStkCoCd and Deleted.OdDmStkNo= DDmStkNo and Deleted.OdDmStkYy= DDmStkYy and Deleted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
   else
   begin
  update DmYyMm set  
    DDmCrWt= DDmCrWt- (select isnull(sum(Del.OdJLWt), 0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm),  
  DDmDrWt= DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del  
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd= DJLRmCd 
   and Deleted.OdJLLotNo= DJLLotNo and Deleted.OdJLLn1= DJLLn1 and Deleted.OdJLLn2= DJLLn2 and Deleted.OdJLLn3 = DJLLn3 and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
 End  
  end
    
	if @mode in ('A','C')
	begin
		Select @row= count('x') from inserted where Inserted.OdTc='JRI' and 
			exists (Select 'a' from Bag where BCoCd=OdDmStkCoCd and BYy=OdDmStkYy and BChr = OdDmStkChr and BNo=OdDmStkNo) and
				not exists 
				(Select 'a' from Bag where 
							BCoCd = OdDmStkCoCd and BOdDmCd=OdDmCd and BOdSfx=OdSfx and BOdDmTcTyp = 'DM' and
							BOdDmSz=OdDmSz and BYy = OdDmStkYy and BChr =OdDmStkChr and BNo = OdDmStkNo and BCls='Y' and BLoc='PFG' and BRepYN='N')
 
		if @row>0 
		begin
			select @errMsg = 'Given DmStkNo not a valid Bag No For this Dsg Code-Sz-Suffix or Bag not closed or Bag not in FG or Bag already in Repair'
			goto ErrHndlr
		end
	end 
 
 If @mode in ('A','C')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'   
 Begin  

  /*** Update InvHd ***/  
  update InvHd set  
  @wInFob = convert (decimal (16, 2), InFOB+ IsNull((Select sum((ins.OdSalPrc*ins.OdOrdQty)+ins.OdJLVchVal) from inserted ins where   
   ins.OdCoCd= InCoCd and ins.OdTc= InTc and ins.OdYy= InYy   
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
  Join OrdMst On OmCoCd=OdCoCd and OmTc=OdTc and OmChr=odChr and OmNo=OdNo and OmIdNo=OdOmIdNo   
  where InCoCd= OdCoCd and InTc= OdTc and InYy= OdYy and InChr= OdChr and InNo= OdNo   
      and (exists (Select 'x' from Param where PTyp= 'TCTYP' and PMCD ='FG' and PValue= OdTc)  
        or   
        exists (Select 'x' from Param where PTyp= 'TCTYP' and   
            PMCD = (Select PValue from Param Where PTyp='TC' and PMCd= OdTc)))  

  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr, DDmStkNo, DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt,DdmStkCoCd)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy,OdDmStkChr,OdDmStkNo,OmFrDmLoc,convert(int,convert(varchar(6),OdOmDt,112)),0,0,@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0,OdDmStkCoCd   
   from inserted join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DCoCd= OdCoCd and DDmCd=OdDmCd   
        and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy= DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmFrDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and OdDmStkCoCd =DdmStkCoCd and 
        DYyMm= convert(int,convert(varchar(6),OdOmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
         
  /* Current Partition */   
  if @wIsJLTc ='N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd =DDmStkCoCd and Ins.OdDmStkNo = DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd =DDmStkCoCd and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdDmCd= DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkCoCd =DDmStkCoCd  
   and Inserted.OdDmStkYy = DDmStkYy and Inserted.OdDmStkChr = DDmStkChr and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdCoCd= DCoCd 
   and OmFrDmLoc=DLoc  and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd= DJLRmCd 
   and Inserted.OdJLLotNo= DJLLotNo and Inserted.OdJLLn1= DJLLn1 and Inserted.OdJLLn2= DJLLn2 and Inserted.OdJLLn3 = DJLLn3  and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 end  
  end
  
 if @mode in ('C','D')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'     
 Begin  
 if @wIsJLTc = 'N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst  
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty-(select isnull(sum(Del.OdOrdQty),0) from Deleted Del Join OrdMst   
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdCoCd= DCoCd and Del.OdDmCd=DDmCd   
   and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and Del.OdDmStkCoCd =DDmStkCoCd and Del.OdDmStkNo= DDmStkNo and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdDmCd= DDmCd   
   and Deleted.OdSfx= DOdSfx and Deleted.OdDmSz= DDmSz and Deleted.OdDmStkCoCd =DDmStkCoCd and Deleted.OdDmStkNo = DDmStkNo and Deleted.OdDmStkYy = DDmStkYy and 
   Deleted.OdDmStkChr = OdDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
  else
  update DmYyMm set  
  DDmCrWt=DDmCrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst  
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst   
   on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd = DJLRmCd and Deleted.OdJLLotNo = DJLLotNo and Deleted.OdJLLn1 = DJLLn1 and
   Deleted.OdJlLn2 = DJLLn2 and Deleted.OdJllN3 = DJLLn3  and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
 If @wTcTyp = 'JT'     
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty- IsNull((Select sum(del.OdOrdQty) from deleted del where   
     (del.OdTc='JIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt- IsNull((Select sum(del.OdJLWt) from deleted del where   
    (del.OdTc='LIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from deleted where   
    OrdDsg.OdCoCd= deleted.OdCoCd and OrdDsg.OdIdNo=deleted.OdBaseIdkey  
   and Exists (Select 'x' from TcMst Where TmCoCd=deleted.OdCoCd and TmTc=deleted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls = (case when @wIsJLTc ='N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
				else (case when OrdDsg.OdJLWt= OrdDsg.OdJLExpWt and OrdDsg.OdJLWt<> 0 then 'Y' else 'N' end)   end)
   from deleted del where OrdDsg.OdCoCd= Del.OdCoCd and OrdDsg.OdIdNo= Del.OdBaseIdKey  
      
   and Exists (Select 'x' from TcMst Where TmCoCd=del.OdCoCd and TmTc=del.OdTC and TmBaseTCs<>'')  
 
  If @mode in ('D','C')
  Begin  
   Update Bag set
		BRepYN = 'N',   
		BRepQty = 0,
		BRepINDt                ='01/01/1980',
		BRepLoc  = '',
		BRepGrWt	= 0,
		BRepOdTc = '',
		BRepOdYy = '',
		BRepOdChr = '',
		BRepOdNo = 0,
		BRepOdSr = 0
   from Deleted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRI'

   Update Bag set
   		BRepYN = 'Y'
   from Deleted
		where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRO'

  End    

 End   
  
 End   
  
 if @mode in ('A','C')  
 begin  
 If @wTcTyp = 'FG' or @wTcTyp = 'JT'    
 Begin  
  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr,DDmStkNo,DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt,DdmStkCoCd)
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy, OdDmStkChr, OdDmStkNo,OmToDmLoc,convert(int,convert(varchar(6),OdOmDt,112)),0,0,
					@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0,OdDmStkCoCd  
     
   from inserted join OrdMst on OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DCoCd= OdCoCd and DDmCd=OdDmCd   
       and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy = DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmToDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and OdDmStkCoCd =DdmStkCoCd and 
       DYyMm= convert(int,convert(varchar(6),OdOmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  
  if @wIsJLTc = 'N'
  update DmYyMm set  
  DDmCrQty=DDmCrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd= DDmStkCoCd and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrQty=DDmDrQty+(select isnull(sum(Ins.OdOrdQty),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdCoCd= DCoCd and Ins.OdDmCd=DDmCd   
   and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and Ins.OdDmStkCoCd= DDmStkCoCd and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkNo= DDmStkNo and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkChr = DDmStkChr and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdDmCd=DDmCd   
   and Inserted.OdSfx= DOdSfx and Inserted.OdDmSz= DDmSz and Inserted.OdDmStkCoCd= DDmStkCoCd and Inserted.OdDmStkNo= DDmStkNo and Inserted.OdDmStkChr = DDmStkChr and Inserted.OdDmStkYy = DDmStkYy 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd = DJLRmCd and Inserted.ODJLLotNo = DJLLotNo and Inserted.ODJLLn1 = DJLLn1 and Inserted.ODJLLn2 = DJLLn2 and 
   Inserted.ODJLLn3 = DJLLn3 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OdOmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 If @wTcTyp = 'JT'   
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty + IsNull((Select sum(Ins.OdOrdQty) from Inserted Ins where   
   (ins.OdTc='JIN' Or Ins.OdCoCd= OrdDsg.OdCoCd)  and Ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt+ IsNull((Select sum(ins.OdJLWt) from Inserted Ins where   
   (ins.OdTc='LIN' Or ins.OdCoCd= OrdDsg.OdCoCd) and ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from Inserted where 
    OrdDsg.OdCoCd= Inserted.OdCoCd and OrdDsg.OdIdNo=Inserted.OdBaseIdkey  
    and Exists (Select 'x' from TcMst Where TmCoCd=Inserted.OdCoCd and TmTc=Inserted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls =(case when @wIsJLTc = 'N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
			else (case when OrdDsg.OdJLWt = OrdDsg.ODJLExpWt and OrdDsg.ODJLWt <> 0 then 'Y' else 'N' end) end)
   from Inserted Ins where (ins.OdTc = 'LIN' or ins.OdTc='JIN' Or  OrdDsg.OdCoCd= Ins.OdCoCd)   
   and OrdDsg.OdIdNo= Ins.OdBaseIdKey and Exists (Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTC and TmBaseTCs<>'')  
  if @mode in ('A','C')  
  Begin   
   Update Bag set   
		BRepYN = 'Y',
		BRepQty = OdOrdEnt,
		BRepINDt = getdate(),
		/* (Select OmDt From OrdMst Where OmCoCd= OdCoCd And OmTc = OdTc   
		And OmYy = OdYy And OmChr= OdChr And OmNo= OdNo And OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo)*/
		BRepLoc  = (Select HDefRepBagLoc from Head where HCoCd = Inserted.OdCoCd),
		BRepGrWt = OdRepGrWt,
		BRepOdTc = OdTc,
		BRepOdYy = OdYy,
		BRepOdChr = OdChr,
		BRepOdNo = OdNo,
		BRepOdSr = OdSr
   from Inserted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRI'

   Update Bag set
		BRepYN= 'N'
   from Inserted where BCoCd= OdDmStkCoCd and BYy= OdDmStkYy and BChr= OdDmStkChr and BNo= OdDmStkNo and OdTc= 'JRO'
  End  
 End  
  
 End    
  
  
if @mode in ('C','A') and @wTcTyp = 'JT'   
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') from Inserted Ins   
   where exists(Select 'x' from OrdDsg B Where  
      B.OdIdNo=Ins.OdBaseIdkey and B.OdOrdQty < B.OdExpQty)  
    and (exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>''))  
  If @Row> 0  
  begin   
   select @errmsg= 'Order Qty Should Be >= Matched Qty for Base Voucher'    
   goto ErrHndlr    
  end
 end
else
begin
 Select @Row=Count('x') from Inserted Ins   
   where 
    exists(Select 'x' from OrdDsg B Where  
      B.OdIdNo=Ins.OdBaseIdkey and cast(B.OdJLWt as decimal(10,3))< cast(B.OdJLExpWt as decimal(10,3)))  
    and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>'')  
 If @Row> 0  
  begin   
   select @errmsg= 'Order Wt Should Be >= Matched Wt for Base Voucher'    
   goto ErrHndlr    
  end
 end   
end    
  
If @mode in ('C','A') and  (@wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then 
	  (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) 
	  else (case when OmFrDmDc = 'C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Ins.OdCoCd= DCoCd and Ins.OdDmCd= DDmCd and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and   
    Ins.OdDmStkCoCd =DdmStkCoCd and Ins.OdDmStkYy = DDmStkYy and Ins.OdDmStkChr = DDmStkChr and Ins.OdDmStkNo = DDmStkNo and 
     (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Desing Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
end
else
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrWt)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.OdJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3=DJLLn3 and
	(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Size'  
     GoTo ErrHndlr  
     end  

end
end  
If @mode in ('D') and  (@wTcTyp = 'FG' or @wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 -- rewrite checking JCN why and remove FG and see what to simplify 
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd = Del.OdCoCd and LocTyp =(Case PValue When 'JT' Then 'S' End)  and   
            LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then (case when OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and   
    Del.OdDmStkCoCd =DdmStkCoCd and Del.OdDmStkYy = DDmStkYy and Del.OdDmStkChr = DDmStkChr and Del.OdDmStkNo = DDmStkNo and 
    (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This Design Code, Suffix, ID'  
     GoTo ErrHndlr  
     end  
	end
else
begin
 -- rewrite checking JCN why and remove FG and see what to simplify 
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Del.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End)  and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrwT)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Del.OdCoCd= DCoCd and Del.ODJLRmCd = DJLRmCd and Del.ODJLLotNo = DJLLotNo and Del.ODJLLn1 = DJLLn1 and Del.ODJLLn2 = DJLLn2 and Del.ODJLLn3 = DJLLn3 and
(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmToDmLoc when OmFrDmDC = 'D' then OmFrDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Sz'  
     GoTo ErrHndlr  
     end  
end

end  
return    
ErrHndlr:    
 begin   select @errmsg='|'+@errmsg    
  raiserror(@errmsg,16,1)   rollback transaction     
 end    
end   



GOGO

ALTER TRIGGER [dbo].[FgRm_INUPDL] ON [dbo].[FgRm] 
       FOR INSERT,UPDATE,DELETE  
		NOT FOR REPLICATION
AS  
BEGIN  
	/* Trigger Version 3.3.0 - Speed Optimization */
	DECLARE @ROW INT  
	DECLARE @mode char(1)  
	DECLARE @errmsg varchar(255)  
	Declare @wTcTyp Varchar(3) 
	Declare @wEntTyp Varchar(8)
	DECLARE @ReplAgt BIT  
	DECLARE @wCurrPrtn VarChar(1) 
	
	IF @@ROWCOUNT= 0   RETURN  
	
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'FgRm' )  
	RETURN  
	
	select @mode=case (select count('x') from inserted) when 0 then 'D'  
	else  
		case (select count('x') from deleted) when 0 then 'A' else 'C' end  
	end   
	
	Select @wCurrPrtn='C'
	if @mode in ('A','C')
	Begin
		select @wTcTyp= max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= FrTc and PSCd=''
	end

	if @mode in ('D')
	Begin
		select @wTcTyp= max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= FrTc and PSCd=''
	end

	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))


	--5717 	/******** Only One TcTyp records can be Inserted at a time.********/	
	select @Row= Count(distinct(PValue)) from inserted Join param On
		PTyp= 'TC' and PMCd= FrTc and PSCd= '' Where FrPrtKey=@wCurrPrtn 
	If @Row > 1
	BEGIN
		Select @ErrMsg= 'Transactions of only One TcType can be Acted On'
	  	GoTo ErrHndlr
	END
	--5717  /******** Only One TcTyp records can be Deleted at a time. *********/	
	select @Row= Count(distinct(PValue)) from deleted Join param on
		PTyp='TC' and PMcd= FrTc and PSCd='' Where FrPrtKey=@wCurrPrtn 
	If @Row > 1
	BEGIN
		Select @ErrMsg='Transactions of only One TcType can be Acted On'
		GoTo ErrHndlr
	END
	--5717
	If @Mode = 'C'
	Begin
		/******* If Mode = 'C' then Cannot change Key Values********/	
		select @Row=Count('x') from Deleted Del where Del.FrPrtKey=@wCurrPrtn and
			 Not Exists (Select 'x' from Inserted Ins where Ins.FrPrtKey=Del.FrPrtKey 
					and Ins.FrIdNo=Del.FrIdNo and Ins.FrCoCd=Del.FrCoCd 
					and Ins.FrTc=Del.FrTc and Ins.FrYy= Del.FrYy
					and Ins.FrChr= Del.FrChr and Ins.FrNo=Del.FrNo
					and Ins.FrSr=Del.FrSr and Ins.FrSrNo= Del.FrSrNo)
		If @Row > 0
		begin		  
			Select @ErrMsg='Cannot Change FgRm Key values'
			GoTo ErrHndlr
		end
	End
	
	/* Zubin 212 Lot No. Shld be blank when HLotNoYN= 'N' in head */ 
	--5717 
	If @wTcTyp in ('FB', 'FWB') 
	Begin
		/* Current Partition */		
		Select @Row= Count('x') From Inserted Ins Where FrPrtKey=@wCurrPrtn and FrLotNo<> '' and 
			Not Exists (Select 'x' From Head where HCoCd= Ins.FrCoCd and HCd= 'ZSELF' and 
				HLotNoYN= 'Y') 
	  	If @Row> 0 
		begin
			Select @ErrMsg= 'Lot No. Should Be Blank As ''Implement Lot No. For D/C/X/G/P/S/L/M'' Option Is ''N'' In Head Master (Tab 5)' 
		  	GoTo ErrHndlr
		end 

	--	--/* Zubin 212 Lot No. For 'D' and 'C' only */ 
	--	--/* Current Partition */	
	--	--/* Manali 3.03 - 16/07/08 - Lot No. open for Acc ('X') also */	
	--	--select @Row= Count('x') from inserted ins where FrLotNo <> '' and FrPrtKey=@wCurrPrtn and
	--	--	Not Exists (select 'x' from RmMst where RmZ= 'N' and RmCd= FrRmCd and RmPrtKey=FrPrtKey and RmCtg IN ('D', 'C', 'X','G','P','S','L','M'))
	-- -- 	If @Row> 0
	--	--begin
	--	--	Select @ErrMsg= 'Lot No. Should Be Entered For Diamonds, Color Stones and Accessories Only' 
	--	--	GoTo ErrHndlr
	--	--end
	End

	--5717
	If @Mode = 'C' and @wTcTyp in ('FB', 'FR')
	Begin
		/**** Cannot edit any field if FdDespQty <> 0 ****/
		/* Zubin 212 (FrLotNo added) */
		/* Current Partition */		
		select @Row= Count('x') from deleted Del
			join Fgd on FdPrtKey=Del.FrPrtKey and FdIdNo=Del.FrFdIdNo and FdCoCd=Del.FrCoCd And FdTc=Del.FrTc 
					And FdYy=Del.FrYy And  FdChr=Del.FrChr And FdNo=Del.FrNo And FdSr=Del.FrSr 
			where FdPrtKey=@wCurrPrtn and FdDespQty<>0 and 
				(not exists(Select 'x' from Inserted Ins where
				Del.FrPrtKey=Ins.FrPrtKey and Del.FrCoCd=Ins.FrCoCd and Del.FrSrno=Ins.FrSrNo and Del.FrFrRmLoc=Ins.FrFrRmLoc
				and Del.FrFrRmDc=Ins.FrFrRmDc and Del.FrRmCd=Ins.FrRmCd and Del.FrLotNo=Ins.FrLotNo 
				and Del.FrRmSz=Ins.FrRmSz and Del.FrRmSz2=Ins.FrRmSz2 and Del.FrRmSz3=Ins.FrRmSz3 and Del.FrRmStkRt=Ins.FrRmStkRt
				and Del.FrRmQty=Ins.FrRmQty and Del.FrRmWt=Ins.FrRmWt 
				and Del.FrToRmLoc=Ins.FrToRmLoc and Del.FrToRmDc=Ins.FrToRmDc))
		If @Row > 0
		begin	
			Select @ErrMsg= 'Cannot Edit as Desp Qty<>0'
			GoTo ErrHndlr
		end
	End

	--5717
	If @Mode = 'C' and @wTcTyp in ('FWB', 'FWR')
	Begin
		/**** Cannot edit any field ****/
		/* Zubin 212 (FrLotNo added) */
		/* Current Partition */		
		select @Row= Count('x') from deleted Del
			join Fgd on FdPrtKey=Del.FrPrtKey and FdIdNo=Del.FrFdIdNo and FdCoCd= Del.FrCoCd And FdTc= Del.FrTc And 
				    FdYy= Del.FrYy And  FdChr= Del.FrChr And
				    FdNo= Del.FrNo And FdSr= Del.FrSr 
			where Del.FrPrtKEy=@wCurrPrtn and not exists(Select 'x' from Inserted Ins where Del.FrPrtKey=Ins.FrPrtKey and 
				Del.FrCoCd=Ins.FrCoCd and Del.FrSrNo= Ins.FrSrNo and Del.FrFrRmLoc= Ins.FrFrRmLoc
				and Del.FrFrRmDc= Ins.FrFrRmDc and Del.FrRmCd= Ins.FrRmCd and Del.FrLotNo= Ins.FrLotNo
				and Del.FrRmSz= Ins.FrRmSz and Del.FrRmSz2= Ins.FrRmSz2 and Del.FrRmSz3= Ins.FrRmSz3 and Del.FrRmStkRt= Ins.FrRmStkRt
				and Del.FrRmQty= Ins.FrRmQty and Del.FrRmWt= Ins.FrRmWt 
				and Del.FrToRmLoc= Ins.FrToRmLoc and Del.FrToRmDc= Ins.FrToRmDc) 
		If @Row > 0
		begin	
			Select @ErrMsg= 'Cannot Edit Any Value'
			GoTo ErrHndlr
		end
	End


	If @Mode in ('A','C')
	Begin
		/******* Valid FrTc,FrYy,FrChr, FrNo, FrSr ********/
		/* Current Partition */		
		Select @Row=Count('x') From Inserted Ins Where FrPrtKey=@wCurrPrtn and
  			Not Exists (Select 'x' From Fgd Where FdPrtKey=Ins.FrPrtKey and FdIdNo=FrFdIdNo and FdCoCd= Ins.FrCoCd And 
				    FdTc= Ins.FrTc And FdYy= Ins.FrYy And  FdChr= Ins.FrChr And FdNo= Ins.FrNo And FdSr= Ins.FrSr)
  		If @Row>0
	 	begin
			Select @ErrMsg='FgRm Voucher No. Does Not Exist'
			GoTo ErrHndlr
	  	end
		--5717
		/***** FrSrNo <> 0 *****/
		/* Current Partition */		
		Select @Row= Count('x') From Inserted Ins Where FrPrtKey=@wCurrPrtn and FrSrNo= 0
		If @Row>0
		begin
	  		Select @ErrMsg='Rm. SrNo Cannot be 0'
		  	GoTo ErrHndlr
		end
		--5717
		Select @Row=Count('x') From Inserted Ins
			Where Ins.FrPrtKey=@wCurrPrtn and Not Exists (Select 'x' From RmMst where 
				RmCd= FrRmCd and RmPrtKey=FrPrtKey and RmZ= 'N')
  	  	If @Row>0
		begin
			Select @ErrMsg='Invalid Rm Code'
		  	GoTo ErrHndlr
		end
		--5717
		/********** FrFrRmDc  Valid 'Dc' *************/
		/* Current Partition */		
		Select @Row=Count('x') From Inserted Ins
			Where Ins.FrPrtKey=@wCurrPrtn and  (Not Exists (Select 'x' From Param where PTyp= 'DC'
				and PMCd = Ins.FrFrRmDc and PScd= '')
				or Not Exists (Select 'x' From Param where PTyp= 'DC'
				and PMCd = Ins.FrToRmDc and PScd= ''))
	  	If @Row>0
		begin
			Select @ErrMsg='Invalid From/ To Rm DC'
		  	GoTo ErrHndlr
		end
		--5717
		/******* FrToRmDc should be Opp of  FrFrRmDc *******/
		/* Current Partition */
	  	Select @Row=Count('x') From Inserted Ins
			Where Ins.FrPrtKey=@wCurrPrtn and ((FrToRmDc= FrFrRmDc And FrFrRmDc<> '')  Or
				(FrFrRmDc<> '' and FrToRmDc= ''))
	  	If @Row>0
		begin
			Select @ErrMsg='Invalid To Rm Dc'
			GoTo ErrHndlr
		end
	End


	If @Mode in ('A','C') and @wTcTyp in ('FB')
	Begin
		--5717
		Select @Row= Count('x') From Inserted Ins 
			Where Ins.FrPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Loc where LocCoCd = FrCoCd 
				and LocTyp= 'F' and (LocCd= FrFrRmLoc or LocCd = FrToRmLoc))
	  	If @Row> 0
		begin
			Select @ErrMsg='Atleast One RM Loc Must Be Of Type ''F'' '
		  	GoTo ErrHndlr
		end
		--5717
		/****** If TcTyp= 'FB' Then FrFrRmLoc or FrToRmLoc Should Have LocTyp= 'W' ********/
		/* Current Partition */
		Select @Row=Count('x') From Inserted Ins 
			Where Ins.FrPrtKEy=@wCurrPrtn and Not Exists (Select 'x' From Loc where LocCoCd = FrCoCd 
				and LocTyp= 'W' and (LocCd= FrFrRmLoc or LocCd = FrToRmLoc))
  	  	If @Row>0
		begin
			Select @ErrMsg='Atleast One RM Loc Must Be Of Type ''W'' '
		  	GoTo ErrHndlr
		end

		Select @Row=Count('x') From Inserted Ins Where FrPrtKey=@wCurrPrtn and (FrRmWt=0 AND FrRmQty = 0)
  	  	If @Row>0
		begin
			Select @ErrMsg='Either Wt and Qty should be non zero for the RM'
		  	GoTo ErrHndlr
		end
		--5717
		/****** If TcTyp= 'FB' Then RM Wt/ Qty Can Be Issued From Locations Of Type 'W' Only ********/
		/* Current Partition */
		Select @Row=Count('x') From Inserted Ins 
			Where FrPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Loc where LocCoCd = FrCoCd and 
				LocTyp= 'W' and ((LocCd= FrFrRmLoc and FrFrRmDc= 'C') 	or
	   					 (LocCd= FrToRmLoc and FrToRmDc= 'C')	
				       		)
   					  )
  	  	If @Row>0
		begin
			Select @ErrMsg= 'RM Wt/ Qty Can Be Issued From Locations Of Type ''W'' Only'
		  	GoTo ErrHndlr
		end
		--5717
		select @Row= Count('x') from inserted Ins
			join Fgd on FdPrtKey=Ins.FrPRtKey and FdIdNo=Ins.FrFdIdNo and FdCoCd=Ins.FrCoCd And FdTc=Ins.FrTc And
				    FdYy=Ins.FrYy And  FdChr=Ins.FrChr And FdNo=Ins.FrNo And FdSr=Ins.FrSr
			Where Ins.FrPrtKEy=@wCurrPrtn and (FdBYy<> '' or FdBChr<> '' or FdBNo<> 0) and 
				(Not Exists (select 'x' from OrdRm where OrPrtKey=FdPrtKey and
				    OrCoCd= FdCoCd And OrTc= FdPrdOdTc And
				    OrYy= FdPrdOdYy And OrChr= FdPrdOdChr And
				    OrNo= FdPrdOdNo And OrSr= FdPrdOdSr And 
				    OrRmCd= FrRmCd))
		If @Row> 0
		begin	
			Select @ErrMsg= 'Rm Code Does Not Exist in Prd. Order'
			GoTo ErrHndlr
		end
		--5717
		/***** FrRmCd to exist in FdExpOrder ******/
		/* Current Partition */
		select @Row= Count('x') from inserted Ins
			join Fgd on FdPrtKey=Ins.FrPrtKey and FdIdNo=Ins.FrFdIdNo and FdCoCd=Ins.FrCoCd And FdTc=Ins.FrTc And 
				    FdYy=Ins.FrYy And  FdChr=Ins.FrChr And
				    FdNo=Ins.FrNo And FdSr=Ins.FrSr
			where not exists(select 'x' from OrdRm where OrPrtKey=FdPrtKey and
				    OrCoCd= FdCoCd And OrTc= FdExpOdTc And
				    OrYy= FdExpOdYy And  OrChr= FdExpOdChr And
				    OrNo= FdExpOdNo And OrSr= FdExpOdSr And 
				    OrRmCd= FrRmCd)
		If @Row> 0
		begin	
			Select @ErrMsg= 'Rm Code Does Not Exist in Exp Order'
			GoTo ErrHndlr
		end

	End


	--If @Mode in ('A','C') and @wTcTyp in ('FR')
	--Begin

	--	/********** If TcTyp= 'FR' Then FrFrRmLoc Valid Loc, LocTyp<>'P', 'W' ************/
	--	/* Current Partition */
	--	Select @Row=Count('x') From Inserted Ins
	--		Where Ins.FrPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Loc where LocCoCd = Ins.FrCoCd and 
	--			LocTyp not in ('P', 'W') and LocCd = Ins.FrFrRmLoc)
 -- 	  	If @Row>0
	--	BEGIN
	--		Select @ErrMsg='Invalid From Rm Loc'
	--	  	GoTo ErrHndlr
	--	END

	--	/****** If TcTyp= 'FR' then FrToRmLoc  Valid Loc, LocTyp<>'P', 'W'  ********/
	--	/* Current Partition */
	--	Select @Row=Count('x') From Inserted Ins
	--		Where Ins.FrPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Loc where LocCoCd = Ins.FrCoCd and 
	--			LocTyp not in ('P', 'W') and LocCd = Ins.FrToRmLoc)
 -- 	  	If @Row>0
	--	BEGIN
	--		Select @ErrMsg='Invalid To Rm Loc'
	--	  	GoTo ErrHndlr
	--	END

	--	/****** If TcTyp= 'FR' Then FrFrRmLoc or FrToRmLoc Should Have LocTyp= 'F' ********/
	--	/* Current Partition */
	--	Select @Row=Count('x') From Inserted Ins 
	--		Where Ins.FrPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Loc where LocCoCd = FrCoCd and 
	--			LocTyp= 'F' and (LocCd= FrFrRmLoc or LocCd = FrToRmLoc))
 -- 	  	If @Row>0
	--	begin
	--		Select @ErrMsg='Atleast One RM Loc Must Be Of Type ''F'' '
	--	  	GoTo ErrHndlr
	--	end

	--	/***** FrRmCd to exist in FdExpOrder ******/
	--	/* Current Partition */
	--	select @Row= Count('x') from inserted Ins
	--		join Fgd on FdPrtKey=Ins.FrPrtKey and FdIdNo=Ins.FrFdIdNo and FdCoCd=Ins.FrCoCd And FdTc=Ins.FrTc And
	--			    FdYy=Ins.FrYy And  FdChr=Ins.FrChr And
	--			    FdNo=Ins.FrNo And FdSr=Ins.FrSr
	--		where Ins.FrPrtKEy=@wCurrPrtn and not exists(select 'x' from OrdRm where 
	--			    OrPrtKey=FdPrtKey and OrCoCd= FdCoCd And OrTc= FdExpOdTc And
	--			    OrYy= FdExpOdYy And  OrChr= FdExpOdChr And
	--			    OrNo= FdExpOdNo And OrSr= FdExpOdSr And 
	--			    OrRmCd= FrRmCd)
	--	If @Row> 0
	--	begin	
	--		Select @ErrMsg= 'Rm Code Does Not Exist in Exp Order'
	--		GoTo ErrHndlr
	--	end
	--End

	--5717
	/*** (Jen 2.13) ***/
	If @Mode in ('A','C') and @wTcTyp in ('FWB')
	Begin
		/****** If TcTyp= 'FWB' Then FrFrRmLoc or FrToRmLoc Should Have LocTyp= 'F' ******/
		/* Current Partition */
		Select @Row= Count('x') From Inserted Ins 
			Where Ins.FrPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Loc where LocCoCd = FrCoCd 
				and LocTyp= 'F' and (LocCd= FrFrRmLoc or LocCd = FrToRmLoc))
	  	If @Row> 0
		begin
			Select @ErrMsg='Atleast One RM Loc Must Be Of Type ''F'' '
		  	GoTo ErrHndlr
		end
		--5717
		/****** If TcTyp= 'FWB' Then FrFrRmLoc or FrToRmLoc Should Have LocTyp= 'W' ******/
		/* Current Partition */
		Select @Row= Count('x') From Inserted Ins 
			Where Ins.FrPrtKey=@wCurrPrtn and Not Exists (Select 'x' From Loc where LocCoCd = FrCoCd 
				and LocTyp= 'W' and (LocCd= FrFrRmLoc or LocCd = FrToRmLoc))
  	  	If @Row>0
		begin
			Select @ErrMsg='Atleast One RM Loc Must Be Of Type ''W'' '
		  	GoTo ErrHndlr
		end
		--5717
		/****** If TcTyp= 'FWB' Then FrRmWT or FrRmQty Should be NonZero ******/
		/* Current Partition */
		Select @Row=Count('x') From Inserted Ins Where Ins.FrPrtKey=@wCurrPrtn and (FrRmWt=0 AND FrRmQty = 0)
  	  	If @Row>0
		begin
			Select @ErrMsg='Either Wt and Qty should be non zero for the RM'
		  	GoTo ErrHndlr
		end
		--5717
		/****** If TcTyp= 'FWB' Then RM Wt/ Qty Can Be Issued From Locations Of Type 'F' Only ******/
		/* Current Partition */
		Select @Row=Count('x') From Inserted Ins 
			Where Ins.FrPrtKey=@wCurrPrtn and (Not Exists (Select 'x' From Loc where LocCoCd = FrCoCd and 
				LocTyp= 'F' and ((LocCd= FrFrRmLoc and FrFrRmDc= 'C') 	or
	   					 (LocCd= FrToRmLoc and FrToRmDc= 'C'))))
  	  	If @Row> 0
		begin
			Select @ErrMsg= 'RM Wt/ Qty Can Be Issued From Locations Of Type ''F'' Only'
		  	GoTo ErrHndlr
		end
		--5717
		/****** If FdBagNo<> 0 FrRmCd to exist in FdPrdOrder ******/
		/* Current Partition */
		select @Row= Count('x') from inserted Ins
			join Fgd on FdPrtKey=Ins.FrPrtkey and FdIdNo=Ins.FrFdIdNo and FdCoCd=Ins.FrCoCd And FdTc=Ins.FrTc And
				    FdYy=Ins.FrYy And  FdChr=Ins.FrChr And FdNo=Ins.FrNo And FdSr=Ins.FrSr
			where FrPrtKey=@wCurrPrtn and (FdBYy<> '' or FdBChr<> '' or FdBNo<> 0) and 
				(not exists(select 'x' from OrdRm where 
				    OrPrtKey=FdPrtKey And OrCoCd= FdCoCd And OrTc= FdPrdOdTc And
				    OrYy= FdPrdOdYy And OrChr= FdPrdOdChr And
				    OrNo= FdPrdOdNo And OrSr= FdPrdOdSr And 
				    OrRmCd= FrRmCd))
		If @Row> 0
		begin	
			Select @ErrMsg= 'Rm Code Does Not Exist in Prd. Order'
			GoTo ErrHndlr
		end
	End
	/*** (Jen 2.13) ***/

	--5717
	If @Mode = 'D'
	Begin
		/**** Cannot delete if FdDespQty <> 0****/
		/* Current Partition */
		Select @Row= Count('x') From deleted Del
			Where Del.FrPrtKey=@wCurrPrtn 
				and (exists(select 'x' from Fgd where FdPrtKey=FrPrtKey and FdIdNo=FrFdIdNo and FdCoCd=FrCoCd And 
								FdTc=FrTc And FdYy=FrYy And  FdChr=FrChr And FdNo=FrNo And FdSr=FrSr and FdDespQty<>0))
		If @Row > 0

		begin	
			Select @ErrMsg= 'Cannot Delete as Desp Qty <> 0'
			GoTo ErrHndlr
		end

	End

	/*** UPDATES ***/

	/*** Update SYyMm from FrFrRmLoc ***/
	/* Zubin 212 (FrLotNo added in SYyMm inserts/updates) */ 
	if @mode in ('C','D')
	Begin
		/* Current Partition */
		update SYyMm set
		SRmCrQty= SRmCrQty- (select isnull(sum(Del.FrRmQty), 0) from Deleted Del 
			Join Fg on FgPrtKey=Del.FrPrtKey and FgCoCd= Del.FrCoCd and FgTc= Del.FrTc and FgYy= Del.FrYy and FgChr= Del.FrChr 
			and FgNo= Del.FrNo where Del.FrFrRmDc= 'C' and Del.FrPrtKey=SPrtKey and Del.FrCoCd= SCoCd and Del.FrRmCd=SRmCd 
			and Del.FrLotNo=SLotNo and Del.FrRmSz=SRmSz and Del.FrRmSz2=SRmSz2 and Del.FrRmSz3=SRmSz3 and Del.FrRmStkRt=SRmStkRt and Del.FrFrRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmCrWt= SRmCrWt-(select isnull(sum(Del.FrRmWt), 0) from Deleted Del 
			Join Fg on FgPrtKey=Del.FrPrtKey and FgCoCd= Del.FrCoCd and FgTc= Del.FrTc and FgYy= Del.FrYy and FgChr= Del.FrChr 
			and FgNo= Del.FrNo where Del.FrFrRmDc= 'C' and Del.FrPrtKey=SPrtKey and Del.FrCoCd= SCoCd and Del.FrRmCd=SRmCd 
			and Del.FrLotNo=SLotNo and Del.FrRmSz=SRmSz and Del.FrRmSz2=SRmSz2 and Del.FrRmSz3=SRmSz3 and Del.FrRmStkRt=SRmStkRt and Del.FrFrRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmDrQty= SRmDrQty-(select isnull(sum(Del.FrRmQty), 0) from Deleted Del 
			Join Fg on FgPrtKey=Del.FrPrtKey and FgCoCd= Del.FrCoCd and FgTc= Del.FrTc and FgYy= Del.FrYy and FgChr= Del.FrChr 
			and FgNo= Del.FrNo where Del.FrFrRmDc= 'D' and Del.FrPrtKey=SPrtKey and Del.FrCoCd= SCoCd and Del.FrRmCd=SRmCd 
			and Del.FrLotNo=SLotNo and Del.FrRmSz=SRmSz and Del.FrRmSz2=SRmSz2 and Del.FrRmSz3=SRmSz3 and Del.FrRmStkRt=SRmStkRt and Del.FrFrRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmDrWt= SRmDrWt-(select isnull(sum(Del.FrRmWt), 0) from Deleted Del 
			Join Fg on FgPrtKey=Del.FrPrtKey and FgCoCd= Del.FrCoCd and FgTc= Del.FrTc and FgYy= Del.FrYy and FgChr= Del.FrChr 
			and FgNo= Del.FrNo where Del.FrFrRmDc= 'D' and Del.FrPrtKey=SPrtKey and Del.FrCoCd= SCoCd and Del.FrRmCd= SRmCd 
			and Del.FrLotNo= SLotNo and Del.FrRmSz=SRmSz and Del.FrRmSz2=SRmSz2 and Del.FrRmSz3=SRmSz3 and Del.FrRmStkRt=SRmStkRt and Del.FrFrRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm) 
		from Deleted Join Fg on FgPrtKey=Deleted.FrPrtKey and FgCoCd= Deleted.FrCoCd and FgTc= Deleted.FrTc and FgYy= Deleted.FrYy 
			and FgChr= Deleted.FrChr and FgNo= Deleted.FrNo 
			where Deleted.FrPrtKey=SPrtKey and Deleted.FrCoCd= SCoCd and Deleted.FrRmCd= SRmCd and Deleted.FrLotNo= SLotNo 
			and Deleted.FrRmSz= SRmSz and Deleted.FrRmSz2= SRmSz2 and Deleted.FrRmSz3= SRmSz3 and Deleted.FrRmStkRt= SRmStkRt 
			and Deleted.FrFrRmLoc= SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))=SYyMm
			and (Deleted.FrFrRmDc='D'  or Deleted.FrFrRmDc='C')
	End


	if @mode in ('A','C')
	Begin
		/* Current Partition */
		insert into SYyMm(SPrtKey,SCoCd,SRmCd,SLotNo,SRmSz,SRmSz2,SRmSz3,SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt) 
			select distinct FrPrtKey,FrCoCd,FrRmCd,FrLotNo,FrRmSz,FrRmSz2,FrRmSz3,FrRmStkRt,FrFrRmLoc,convert(int,convert(varchar(6),Fgdt,112)),0,0,0,0
			from inserted join Fg on FgPRtKey=FrPrtKey and FgCoCd=FrCoCd and FgTc=FrTc and FgYy=FrYy and FgChr=FrChr and FgNo=FrNo
			where not exists(select 'x' from  SYyMm where SPrtKey=FrPrtKey and SCoCd= FrCoCd and SRmCd=FrRmCd and SLotNo=FrLotNo 
			and SRmSZ=FrRmSz and SRmSZ2=FrRmSz2 and SRmSZ3=FrRmSz3 and SRmStkRt=FrRmStkRt and SLoc=FrFrRmLoc and
			SYyMm= convert(int,convert(varchar(6),Fgdt,112)))

		/* Current Partition */
		update SYyMm set
		SRmCrQty= SRmCrQty+(select isnull(sum(Ins.FrRmQty), 0) from Inserted Ins 
			Join Fg on FgPrtKey=Ins.FrPrtKey and FgCoCd= Ins.FrCoCd and FgTc= Ins.FrTc and FgYy= Ins.FrYy and FgChr= Ins.FrChr 
			and FgNo= Ins.FrNo where Ins.FrFrRmDc= 'C' and Ins.FrPrtKey=SPrtKey and Ins.FrCoCd= SCoCd and Ins.FrRmCd= SRmCd 
			and Ins.FrLotNo= SLotNo and Ins.FrRmSz=SRmSz and Ins.FrRmStkRt=SRmStkRt and Ins.FrFrRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmCrWt=SRmCrWt+(select isnull(sum(Ins.FrRmWt), 0) from Inserted Ins 
			Join Fg on FgPrtKey=Ins.FrPrtKey and FgCoCd= Ins.FrCoCd and FgTc= Ins.FrTc and FgYy= Ins.FrYy and FgChr= Ins.FrChr 
			and FgNo= Ins.FrNo where Ins.FrFrRmDc= 'C' and Ins.FrPrtKey=SPrtKey and Ins.FrCoCd= SCoCd and Ins.FrRmCd=SRmCd 
			and Ins.FrLotNo=SLotNo and Ins.FrRmSz=SRmSz and Ins.FrRmSz2=SRmSz2 and Ins.FrRmSz3=SRmSz3 and Ins.FrRmStkRt=SRmStkRt and Ins.FrFrRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmDrQty=SRmDrQty+(select isnull(sum(Ins.FrRmQty), 0) from Inserted Ins 
			Join Fg on FgPrtKey=Ins.FrPrtKey and FgCoCd= Ins.FrCoCd and FgTc= Ins.FrTc and FgYy= Ins.FrYy and FgChr= Ins.FrChr 
			and FgNo= Ins.FrNo where Ins.FrFrRmDc= 'D' and Ins.FrPrtKey=SPrtKey and Ins.FrCoCd= SCoCd and Ins.FrRmCd=SRmCd 
			and Ins.FrLotNo=SLotNo and Ins.FrRmSz=SRmSz and Ins.FrRmSz2=SRmSz2 and Ins.FrRmSz3=SRmSz3 and Ins.FrRmStkRt=SRmStkRt and Ins.FrFrRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmDrWt=SRmDrWt+(select isnull(sum(Ins.FrRmWt), 0) from Inserted Ins 
			Join Fg on FgPrtKey=Ins.FrPrtKey and FgCoCd= Ins.FrCoCd and FgTc= Ins.FrTc and FgYy= Ins.FrYy and FgChr= Ins.FrChr 
			and FgNo= Ins.FrNo where Ins.FrFrRmDc= 'D' and Ins.FrPrtKey=SPrtKey and Ins.FrCoCd= SCoCd and Ins.FrRmCd=SRmCd 
			and Ins.FrLotNo=SLotNo and Ins.FrRmSz=SRmSz and Ins.FrRmSz2=SRmSz2 and Ins.FrRmSz3=SRmSz3 and Ins.FrRmStkRt=SRmStkRt and Ins.FrFrRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm)
		from Inserted Join Fg on FgPrtKey=Inserted.FrPrtKey and FgCoCd= Inserted.FrCoCd and FgTc= Inserted.FrTc and FgYy= Inserted.FrYy 
			and FgChr= Inserted.FrChr and FgNo= Inserted.FrNo 
			where Inserted.FrPrtKey=SPrtKey and Inserted.FrCoCd= SCoCd and Inserted.FrRmCd= SRmCd and Inserted.FrLotNo= SLotNo 
			and Inserted.FrRmSz= SRmSz and Inserted.FrRmSz2= SRmSz2 and Inserted.FrRmSz3= SRmSz3 and Inserted.FrRmStkRt= SRmStkRt and Inserted.FrFrRmLoc= SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))=SYyMm
			and (Inserted.FrFrRmDc='D'  or Inserted.FrFrRmDc='C')
	End

	/*** Update SYyMm from FrToRmLoc ***/
	if @mode in ('C','D')
	Begin
		/* Current Partition */
		update SYyMm set
		SRmCrQty= SRmCrQty- (select isnull(sum(Del.FrRmQty), 0) from Deleted Del 
			Join Fg on FgPrtKey=Del.FrPrtKey and FgCoCd= Del.FrCoCd and FgTc= Del.FrTc and FgYy= Del.FrYy and FgChr= Del.FrChr 
			and FgNo= Del.FrNo where Del.FrToRmDc= 'C' and Del.FrPrtKey=SPrtKey and Del.FrCoCd= SCoCd and Del.FrRmCd=SRmCd 
			and Del.FrLotNo=SLotNo and Del.FrRmSz=SRmSz and Del.FrRmSz2=SRmSz2 and Del.FrRmSz3=SRmSz3 and Del.FrRmStkRt=SRmStkRt and Del.FrToRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmCrWt= SRmCrWt-(select isnull(sum(Del.FrRmWt), 0) from Deleted Del 
			Join Fg on FgPrtKey=Del.FrPrtKey and FgCoCd= Del.FrCoCd and FgTc= Del.FrTc and FgYy= Del.FrYy and FgChr= Del.FrChr 
			and FgNo= Del.FrNo where Del.FrToRmDc= 'C' and Del.FrPrtKey=SPrtKey and Del.FrCoCd= SCoCd and Del.FrRmCd=SRmCd 
			and Del.FrLotNo=SLotNo and Del.FrRmSz=SRmSz and Del.FrRmSz2=SRmSz2 and Del.FrRmSz3=SRmSz3 and Del.FrRmStkRt=SRmStkRt and Del.FrToRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmDrQty= SRmDrQty-(select isnull(sum(Del.FrRmQty), 0) from Deleted Del 
			Join Fg on FgPrtKey=Del.FrPrtKey and FgCoCd= Del.FrCoCd and FgTc= Del.FrTc and FgYy= Del.FrYy and FgChr= Del.FrChr 
			and FgNo= Del.FrNo where Del.FrToRmDc= 'D' and Del.FrPrtKey=SPrtKey and Del.FrCoCd= SCoCd and Del.FrRmCd=SRmCd 
			and Del.FrLotNo=SLotNo and Del.FrRmSz=SRmSz and Del.FrRmSz2=SRmSz2 and Del.FrRmSz3=SRmSz3 and Del.FrRmStkRt=SRmStkRt and Del.FrToRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmDrWt= SRmDrWt-(select isnull(sum(Del.FrRmWt), 0) from Deleted Del 
			Join Fg on FgPrtKey=Del.FrPrtKey and FgCoCd= Del.FrCoCd and FgTc= Del.FrTc and FgYy= Del.FrYy and FgChr= Del.FrChr 
			and FgNo= Del.FrNo where Del.FrToRmDc='D' and Del.FrPrtKey=SPrtKey and Del.FrCoCd= SCoCd and Del.FrRmCd=SRmCd 
			and Del.FrLotNo=SLotNo and Del.FrRmSz=SRmSz and Del.FrRmSz2=SRmSz2 and Del.FrRmSz3=SRmSz3 and Del.FrRmStkRt=SRmStkRt and Del.FrToRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm) 
		from Deleted Join Fg on FgPrtKey=Deleted.FrPrtKey and FgCoCd= Deleted.FrCoCd and FgTc= Deleted.FrTc and 
			FgYy= Deleted.FrYy and FgChr= Deleted.FrChr and FgNo= Deleted.FrNo 
			Where Deleted.FrPrtKey=SPrtKey and Deleted.FrCoCd= SCoCd and Deleted.FrRmCd= SRmCd and Deleted.FrLotNo= SLotNo 
			and Deleted.FrRmSz= SRmSz and Deleted.FrRmSz2= SRmSz2 and Deleted.FrRmSz3= SRmSz3 and Deleted.FrRmStkRt= SRmStkRt and Deleted.FrToRmLoc= SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))=SYyMm
			and (Deleted.FrToRmDc='D'  or Deleted.FrToRmDc='C')
	End

	if @mode in ('A','C')
	Begin
		/* Current Partition */
		insert into SYyMm(SPrtKey,SCoCd,SRmCd,SLotNo,SRmSz,SRmSz2,SRmSz3,SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt) 
			select distinct FrPrtKey,FrCoCd,FrRmCd,FrLotNo,FrRmSz,FrRmSz2,FrRmSz3,FrRmStkRt,FrToRmLoc,convert(int,convert(varchar(6),Fgdt,112)),0,0,0,0
			from inserted join Fg on FgPrtKey=FrPrtKey and FgCoCd= FrCoCd and FgTc=FrTc and FgYy=FrYy and FgChr=FrChr and FgNo=FrNo
			where not exists(select 'x' from  SYyMm where SPrtKey=FrPrtKey and SCoCd= FrCoCd and SRmCd=FrRmCd and 
			SLotNo=FrLotNo and SRmSZ= FrRmSz and SRmSZ2= FrRmSz2 and SRmSZ3= FrRmSz3 and SRmStkRt=FrRmStkRt and SLoc=FrToRmLoc and
			SYyMm= convert(int,convert(varchar(6),Fgdt,112)))

		/* Current Partition */
		update SYyMm set
		SRmCrQty= SRmCrQty+(select isnull(sum(Ins.FrRmQty), 0) from Inserted Ins 
			Join Fg on FgPrtKey=Ins.FrPrtKey and FgCoCd= Ins.FrCoCd and FgTc= Ins.FrTc and FgYy= Ins.FrYy and FgChr= Ins.FrChr 
			and FgNo= Ins.FrNo where Ins.FrToRmDc= 'C' and Ins.FrPrtKey=SPrtKey and Ins.FrCoCd= SCoCd and Ins.FrRmCd=SRmCd 
			and Ins.FrLotNo=SLotNo and Ins.FrRmSz=SRmSz and Ins.FrRmSz2=SRmSz2 and Ins.FrRmSz3=SRmSz3 and Ins.FrRmStkRt=SRmStkRt and Ins.FrToRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmCrWt=SRmCrWt+(select isnull(sum(Ins.FrRmWt), 0) from Inserted Ins 
			Join Fg on FgPrtKey=Ins.FrPrtKey and FgCoCd= Ins.FrCoCd and FgTc= Ins.FrTc and FgYy= Ins.FrYy and FgChr= Ins.FrChr 
			and FgNo= Ins.FrNo where Ins.FrToRmDc= 'C' and Ins.FrPrtKey=SPrtKey and Ins.FrCoCd= SCoCd and Ins.FrRmCd=SRmCd 
			and Ins.FrLotNo=SLotNo and Ins.FrRmSz=SRmSz and Ins.FrRmSz2=SRmSz2 and Ins.FrRmSz3=SRmSz3 and Ins.FrRmStkRt=SRmStkRt and Ins.FrToRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmDrQty=SRmDrQty+(select isnull(sum(Ins.FrRmQty), 0) from Inserted Ins 
			Join Fg on FgPrtKey=Ins.FrPrtKey and FgCoCd= Ins.FrCoCd and FgTc= Ins.FrTc and FgYy= Ins.FrYy and FgChr= Ins.FrChr 
			and FgNo= Ins.FrNo where Ins.FrToRmDc= 'D' and Ins.FrPrtKey=SPrtKey and Ins.FrCoCd= SCoCd and Ins.FrRmCd=SRmCd 
			and Ins.FrLotNo=SLotNo and Ins.FrRmSz=SRmSz and Ins.FrRmSz2=SRmSz2 and Ins.FrRmSz3=SRmSz3 and Ins.FrRmStkRt=SRmStkRt and Ins.FrToRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm),
		SRmDrWt=SRmDrWt+(select isnull(sum(Ins.FrRmWt), 0) from Inserted Ins 
			Join Fg on FgPrtKey=Ins.FrPrtKey and FgCoCd= Ins.FrCoCd and FgTc= Ins.FrTc and FgYy= Ins.FrYy and FgChr= Ins.FrChr 
			and FgNo= Ins.FrNo where Ins.FrToRmDc= 'D' and Ins.FrPRtKey=SPrtKEy and Ins.FrCoCd= SCoCd and Ins.FrRmCd=SRmCd 
			and Ins.FrLotNo=SLotNo and Ins.FrRmSz=SRmSz and Ins.FrRmSz2=SRmSz2 and Ins.FrRmSz3=SRmSz3 and Ins.FrRmStkRt=SRmStkRt and Ins.FrToRmLoc=SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))= SYyMm)
		from Inserted Join Fg on FgPrtKey=Inserted.FrPrtKey and FgCoCd= Inserted.FrCoCd and FgTc= Inserted.FrTc and FgYy= Inserted.FrYy 
			and FgChr= Inserted.FrChr and FgNo= Inserted.FrNo 
			where Inserted.FrPrtKey=SPrtKey and Inserted.FrCoCd= SCoCd and Inserted.FrRmCd= SRmCd and Inserted.FrLotNo= SLotNo 
			and Inserted.FrRmSz= SRmSz and Inserted.FrRmSz2= SRmSz2 and Inserted.FrRmSz3= SRmSz3 and Inserted.FrRmStkRt= SRmStkRt and Inserted.FrToRmLoc= SLoc 
			and convert(int,convert(varchar(6),Fgdt,112))=SYyMm
			and  (Inserted.FrToRmDc='D'  or Inserted.FrToRmDc='C')
	End
Return
ErrHndlr:
	BEGIN
		Select @ErrMsg='|'+@ErrMsg
		RaisError(@ErrMsg,16,1)
		Rollback Transaction
	END
END

GOgo

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED



