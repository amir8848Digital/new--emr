Update TcStr set TsStyle = 'B'  where TsStyp = 'CMD' and TsIdName = 'CmdStkCpy' and TsTbl = 'Ord' 
and TsTyp in ('JPT') 

/***** MV Entry ***********************************************************************/
Update tcstr set tsstyle = 'V' where tstbl = 'TxnRm' and TsTyp in ('MV') and TsIdName in ('FraPurHd' ,'FraPurFot','FraVS','FraCnvTot')

update tcstr set  TcStr.TsStyle = tcStrOfPur.TsStyle, TsBlank = tcStrOfPur.TsBlank, TsLbl = TcStrOfPur.TsLbl 
from 
	tcStr 
join tcStr tcStrOfPur 
on
	tcStr.TsTbl  = tcStrOfPur.TsTbl 
	and TcStrOfPur.TsTyp  = 'PR'
	and TcStr.TsIdName = TcStrOfPur.TsIdName
	and TcStr.TsSTyp = TcStrOfPur.TsSTyp
where TcStr.tstyp = 'MV' and TcStr.tsidname in ('TDPURRT' ,'TDPURAMT','TDTORMLOC','WTOTQTY','WTOTWT','TSUPPCD','TCURRCD','TCNVRT','TMULDIV','TMULDIVRS','WCMNAME','TSUPPPAYTERMS','TCNVRTRS', 'TBILLNO','TBOENO','TBILLDT','TBOEDT','TNETAMT','TADDAMT1','TADDAMT2','TADDAMT3','TADDAMT4','TTOTAMT','LblTAddAmt1','LblTAddAmt2','LblTAddAmt3','LblTAddAmt4','LblTotQty','LblTotWt')

/************END ***/



/****** Object:  Trigger [dbo].[Txn_INUPDL]    Script Date: 7/30/2015 8:51:29 AM ******/
SET ANSI_NULLS ON
GOGO
SET QUOTED_IDENTIFIER ON
GOGO


	ALTER TRIGGER [dbo].[Txn_INUPDL] ON [dbo].[Txn] 
	   FOR INSERT,UPDATE,DELETE NOT FOR REPLICATION

	AS
	BEGIN
	/* Trigger Version 3.10.0  */
	DECLARE @ROW INT
	DECLARE @mode char(1)
	DECLARE @errmsg varchar(255)
	DECLARE @wTcTyp varchar(3)
	DECLARE @wEntTyp varchar(8)
	Declare @wHCd Varchar(8) 
	DECLARE @ReplAgt BIT  
	DECLARE @wCurrPrtn VarChar(1)
		
	IF @@ROWCOUNT=0   RETURN
	
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'Txn' )
	RETURN
	
	select @mode=case (select count('x') from inserted) when 0 then 'D'
	else
		case (select count('x') from deleted) when 0 then 'A' else 'C' end
	end

	Select @wHCd= 'ZSELF'
	Select @wCurrPrtn='C'

	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))
	
	 If @ReplAgt = 1         Return
      

	/******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */
	if @mode='C'
	begin
		Select @row=count('x') From Inserted Where TPrtKey<>@wCurrPrtn
		if @row> 0 
		begin
			select @errmsg= 'Cannot Edit Records of Previous Partition'
			goto ErrHndlr
		end

		Select @row=count('x') From Deleted del
					 Where Exists(Select 'x' From Inserted ins Where ins.TIdNo=del.TIdNo and ins.TCoCd=del.TCoCd and ins.TTc=del.TTc and
									ins.TYy=del.TYy and ins.TChr=del.TChr and ins.TNo=del.TNo and ins.TPrtKey<>del.TPrtKey) 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Edit Partition Key Field'
			goto ErrHndlr
		end

	End

	if @mode='A'
	Begin
		Select @row=count('x') From Inserted Where TPrtKey <> @wCurrPrtn 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Insert Partition Key Other Than ''C'''
			goto ErrHndlr
		end
	end 

	If @Mode='D' 
	Begin 
		select @row= count('x') from deleted del where del.TPrtKey<>@wCurrPrtn
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
		select @ROW =Count('x')  From inserted where TPrtKey=@wCurrPrtn and 
		Exists ( select 'x' From SysDb where (SysId='EMR' and (SysDb='LTP' Or ','+sdCocd+',' Not  like '%,'+ TCoCd +',%')) 
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
		select @ROW =Count('x')  From Deleted where TPrtKey=@wCurrPrtn and 
		Exists ( select 'x' From SysDb where (SysId='EMR' and (SysDb='LTP' Or ','+sdCocd+',' Not  like '%,'+ TCoCd +',%'))
			or (SysId='CRM' and SysDb='WEB'))
		if @Row > 0		
			Begin 
				select @errmsg= 'Cannot Delete ' 
				goto ErrHndlr 
			end 
		End 
	End

	/******** Only One TcTyp records can be Inserted at a time.********/
	/* Partition Key Not required */
	select @Row= Count(distinct(PValue)) from inserted Join param On
		PTyp= 'TC' and PMCd= TTc and PScd= ''
	If @Row > 1
	BEGIN
		Select @ErrMsg= 'Transactions of only One TcType can be Acted On'
	  	GoTo ErrHndlr
	END

	/******** Only One TcTyp records can be Deleted at a time. *********/
	/* Partition Key Not Required */
	select @Row= Count(distinct(PValue)) from deleted Join param on
		PTyp='TC' and PMcd=TTc and PSCd=''
	If @Row > 1
	BEGIN
		Select @ErrMsg='Transactions of only One TcType can be Acted on'
		GoTo ErrHndlr
	END

	/*** Storing Tc Type In The Variable @wTcTyp ***/
	if @mode in ('A','C')
	Begin
		select @wTcTyp=max(PValue) from inserted join Param on PTyp= 'TC' and PMCd= TTc and PSCd=''
	end

	if @mode in ('D')
	Begin
		select @wTcTyp=max(PValue) from deleted join Param on PTyp= 'TC' and PMCd= TTc and PSCd=''
	end

	/*** Storing Entry Type In The Variable @wEntTyp ***/
	select @wEntTyp= PValue1 from Param  where PTyp= 'TCTYP' and PMCd=@wTcTyp and PSCd=''

	if @mode in ('A','C')
	Begin
		/*** Transactions Of Only 'TxnBag' and 'TxnRm' Type Permitted ***/
		/* Current Partition */
		select @row= count('x') from inserted where TPrtKey=@wCurrPrtn and not exists 
			(select 'x' from Param where PTyp= 'TC' and PMCd= TTc and 
				PValue in (Select PMCd from Param where PTyp= 'TCTYP' 
						and PValue1 in('TXNBAG', 'TXNRM', 'TZ'))
				   	   )
		if @row> 0 
		begin
   			select @errmsg= 'Transactions Of Only ''Txn''/ ''Tz'' Types Permitted'
			goto ErrHndlr
		end

		/******  TCoCd = HCoCd   *******/
		/* Current Partition */
		select @row=count('x') from inserted where TPrtKey=@wCurrPrtn and 
					not exists (select 'x' from Head where HCoCd= TCoCd and HCd= @wHCd)
		if @row>0 
		begin 
			select @errMsg= 'Company Code Not Valid'
			goto ErrHndlr
		end

		/* Current Partition */
		select @row= count('x') from inserted where TPrtKey=@wCurrPrtn and not exists 
			(select 'x' from Param where PTyp='YY' and PMCd= TYy)
		if @row> 0 
		begin
			select @errmsg= 'Voucher Year Not Defined'
			goto ErrHndlr
		end

		/*****4.1.2********/	
		Select @Row=Count('x') From Inserted ins where 
			(Select hHolChkYN From Head where HCd='ZSELF' and HCoCd=TCoCd)='Y' and
			(@Mode='A' or (@Mode='C' and  ins.TDt<>IsNull((Select TDt From Deleted del where 
									ins.TPrtKey=del.TPrtKey and ins.TIdNo=del.TIdNo 
									and ins.TCoCd= del.TCoCd and ins.TTc= del.TTc and ins.TYy= del.TYy 
										and ins.TChr= del.TChr and ins.TNo= del.TNo),ins.TDt))) and  
							 exists(Select 'x' From DtTbl Where DtTag='HOL' and TDt=DtDt and DtHolYN='Y')
		If @Row> 0
		begin  
			select @errmsg= 'Cannot Add Voucher On Holidays.'  
			goto ErrHndlr  
		end  
		
		

		/***************************** urmi Jul 03 Ch74 *****************************************/
		/* Current Partition */
		select @row= count('x') from inserted Where @wTcTyp in ('TZ') and TPrtKey=@wCurrPrtn 
			and not exists(select 'x' from Param where PTyp= 'ImpTyp' and PMCd= TImpTyp) 
		if @row> 0 
		begin
			select @errmsg= 'Import Type Not Defined'
			goto ErrHndlr
		end

		/* Current Partition */
		select @row= count('x') from inserted where @wTcTyp not in ('TZ') and TPrtKey=@wCurrPrtn 
				and TImpTyp<> ''  
		if @row> 0 
		begin
			select @errmsg= 'Import Type Should Not Be Defined'
			goto ErrHndlr
		end
		/***************************** urmi Jun 03 Ch74 *****************************************/


		/**** Manali 3.8.0 - TIMPCTG  *****/
		select @row= count('x') from inserted Where @wTcTyp in ('TZ') and TPrtKey=@wCurrPrtn 
			and not exists(select 'x' from Param where PTyp= 'ImpCtg' and PMCd= TImpCtg) 
		if @row> 0 
		begin
			select @errmsg= 'Import Category Not Defined'
			goto ErrHndlr
		end

		select @row= count('x') from inserted Ins Where @wTcTyp in ('TZ') and TPrtKey=@wCurrPrtn 
			and Exists(Select 'x' from Deleted Del Where Del.TCoCd=Ins.TCoCd and Del.TTc=Ins.TTc and Del.TYy=Ins.TYy and Del.TChr=Ins.TChr and Del.TNo=Ins.TNo and Del.TPrtKey=Ins.TPrtKey and Del.TIdNo=Ins.TIdNo
										and Del.TImpCtg<>Ins.TImpCtg and 
										exists(Select 'x' from Txndz Where 
										TzCoCd=Del.TCoCd and TzTc=Del.TTc and TzYy=Del.TYy and TzChr=Del.TChr and TzNo=Del.TNo and TzPrtKey=Del.TPrtKey and TzTIdNo=Del.TIdNo
										and Convert(Decimal(14,3), TzRmDespWt)>0))
		if @row> 0 
		begin
			select @errmsg= 'Import Category Cannot Be Changed After Matching Is Done'
			goto ErrHndlr
		end


		select @row= count('x') from inserted where @wTcTyp not in ('TZ') and TPrtKey=@wCurrPrtn 
				and TImpCtg<> ''  
		if @row> 0 
		begin
			select @errmsg= 'Import Category Should Not Be Defined'
			goto ErrHndlr
		end
		/**** Manali 3.8.0 - TIMPCTG  *****/

		/*** (Jen 3.1.1) ***/
		/* Current Partition */		
		select @row= count('x') from inserted where TPRtKey=@wCurrPrtn and @wTcTyp in ('TXNBAG', 'TXNRM') and TCmCd<> ''  
		if @row> 0 
		begin
			select @errmsg= 'Customer Code Should Not Be Defined' 
			goto ErrHndlr
		end

		/* Current Partition */
		select @row= count('x') from inserted where TPrtKey=@wCurrPrtn and @wTcTyp in ('TZ') and TCmCd<> '' 
			and not exists(select 'x' from CustMst where CmCtg= 'C' and CmCd= TCmCd And CmValidYN = 'Y')
		if @row> 0 
		begin
			select @errmsg= 'Customer Code Not Defined' 
			goto ErrHndlr
		end
		/*** (Jen 3.1.1) ***/

		/* Current Partition */
		select @row= count('x') from inserted where TPRtKey=@wCurrPrtn and not exists 
			(select 'x' from vParam where vPCoCd= TCoCd and vPTyp= 'CHR' and 
				vPMCd= TTc and vPSCd= TChr)
  		if @row> 0  
		begin
			select @errmsg= 'Voucher Character Not Defined'
	   		goto ErrHndlr
  		end

		/* Current Partition */
	  	select @row= count('x') from inserted where TPrtKey=@wCurrPrtn and TNo= 0 
	  	if @row> 0 
  		begin
	   		select @errmsg= 'Voucher Number Cannot Be Zero'
   			goto ErrHndlr
	  	end

		/*** (Jen 3.2.0) ***/
		select @row= count('x') from inserted where TDt in('01/01/1980', '01/01/1900')
		if @row> 0 
		begin
			select @errmsg= 'Transaction Date Cannot Be Blank'
			goto ErrHndlr
		end
		/*** (Jen 3.2.0) ***/

		/* Zubin 213 */ 
		/* Current Partition */
	  	select @row= count('x') from inserted where @wTcTyp in ('TZ', 'PR') and TPrtKey=@wCurrPrtn and
			Convert(Decimal(16, 4), TCnvRtRs)<= 0 
  		if @row> 0 
	  	begin
  			select @errmsg= 'Rupees Conversion Rate Should Be > 0'
	   		goto ErrHndlr
  		end

		/* Current Partition */
		select @row= count('x') from inserted where @wTcTyp in ('TZ', 'PR') AND TPrtKey=@wCurrPrtn and
			not exists (select 'x' from Param where PTyp= 'MULDIV' and PMCd= TMulDivRs and PSCd= '')
		if @row> 0 
		begin
			select @errmsg= 'Rupees Multiplication Factor Not Defined' 
			goto ErrHndlr
		end

		/* Zubin 213 */ 
		/******  If @wEntTyp = 'Tz' then TUseBefDt  <>  '01/01/1980'  or  '01/01/1900'  (The dates '01/01/1980'  and  '01/01/1900'  are considered as blank) ******/
		/* Current Partition */
	  	select @row= count('x') from inserted where TPrtKey=@wCurrPrtn and
					@wEntTyp in ('Tz') and TUseBefDt in('01/01/1980', '01/01/1900') 
	  	if @row>0 
  		begin
	   		select @errmsg= 'Use Before Date Cannot Be Blank'
   			goto ErrHndlr
	  	end

		/* Current Partition */
		select @row= count('x') from inserted where TPRtKey=@wCurrPrtn and TFrRmLoc<> '' and 
			not exists (select 'x' from Loc where 
					LocCoCd= TCoCd and LocTyp<> 'P' and LocCd= TFrRmLoc)
  		if @row>0 
	  	begin
   			select @errmsg= 'Default From Rm Loc Not Defined'
   			goto ErrHndlr
	  	end

/*  ######################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ######################  */
		/* Current Partition */
		select @row= count('x') from inserted Ins where Ins.TPrtKEy=@wCurrPrtn and
			Ins.TToRmLoc <> ISNULL((SELECT Del.TToRmLoc FROM Deleted Del WHERE Ins.TIdNo=Del.TIdNo And Ins.TCoCd = Del.TCoCd And
				Ins.TTc = Del.TTc AND Ins.TYy = Del.TYy AND Ins.TChr = Del.TChr AND Ins.TNo = Del.TNo and Ins.TPrtKEy=Del.TPrtKey), '')
			AND TToRmLoc <> ''
			AND	NOT EXISTS (SELECT 'X' FROM Loc WHERE LocCoCd= Ins.TCoCd AND
				 LocTyp <> 'P' AND LocCd = Ins.TToRmLoc AND LocValidYN = 'Y')
  		if @row>0
	  	begin
   			select @errmsg= 'Default To Rm Loc Not Defined/Invalid'
   			goto ErrHndlr
	  	end

/* Bef 2.06
		select @row= count('x') from inserted where TToRmLoc<> '' and
			not exists (select 'x' from Loc where
					LocCoCd= TCoCd and LocTyp<> 'P' and LocCd= TToRmLoc And LocValidYN = 'Y')
  		if @row>0
	  	begin
   			select @errmsg= 'Default To Rm Loc Not Defined'
   			goto ErrHndlr
	  	end
*/
/*  #########################################  Manoj  #########################################  */

		/* Current Partition */
		select @row= count('x') from inserted Ins where Ins.TPrtKey=@wCurrPRtn and Ins.TFrRmLoc = Ins.TToRmLoc 
			 and (Ins.TFrRmLoc <> '' or Ins.TToRmLoc <> '')
  		if @row>0 
	  	begin
   			select @errmsg= 'To And From Rm Locations Cannot Be Same'
   			goto ErrHndlr
	  	end

		/* Current Partition */
	  	select @row= count('x') from inserted where TPrtKey=@wCurrPrtn and TFrBLoc<> '' and 
			not exists (select 'x' from Loc where 
					LocCoCd= TCoCd and LocTyp= 'P' and LocCd= TFrBLoc)
	  	if @row>0 
  		begin
   			select @errmsg= 'Default From Bag Loc Not Defined'
   			goto ErrHndlr
	  	end
		/* **** Zubin 211 **** */ 

		/* Current Partition */
	  	select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and TPrtKey=@wCurrPrtn and
			@wTcTyp IN ('DT', 'BV', 'BS', 'BM', 'MLT') and TFrBLoc= '' and 
			IsNull((Select hOthLocBag from Head where HCoCd= TCoCd and HCd= @wHCd), '') = 'N'
	  	if @row>0 
		begin
	   		select @errmsg= 'Default From Bag Location Cannot Be Blank'
			goto ErrHndlr
	  	end
/* **** Zubin 211 **** */ 

/* Zubin 212 */ 
		/* Current Partition */
	  	select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and TPrtKey=@wCurrPrtn and
			@wTcTyp = 'DT' and TFrRmLoc<> (select LocCd from Loc where LocCoCd= TCoCd and LocTyp= 'W') AND 
				TToRmLoc<> (select LocCd from Loc where LocCoCd= TCoCd and LocTyp= 'W') 
	  	if @row> 0 
		begin
	   		select @errmsg= 'Either RmLocFr Or RmLocTo Should Belong To ''W'' Loc Type' 
			goto ErrHndlr
	  	end
/* Zubin 212 */ 

/* **** Zubin 211 **** */ 
		/* Current Partition */
	  	select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and TPrtKey=@wCurrPRtn and
			@wTcTyp IN ('BM') and Convert(Decimal(16, 1), TTrfQty)> 0 
	  	if @row>0 
		begin
	   		select @errmsg= 'Trf Qty Should Be Zero'
			goto ErrHndlr
	  	end
		/* **** Zubin 211 **** */ 

			/***** Manali 3.8.0 - Multiple Issue/Receipt In BS  ****/
	  	select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and TPrtKey=@wCurrPRtn and
			@wTcTyp IN ('BS') and  Convert(Decimal(16, 1), TTrfQty)= 0
			and Not Exists(Select 'x' from param Where PTyp='IR' and PMCd=TMultiIr)
	  	if @row>0 
			begin
	   		select @errmsg= 'Invalid Multi I/R'
			goto ErrHndlr
	  	end

	  	select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and TPrtKey=@wCurrPRtn and
			@wTcTyp IN ('BS') and  Convert(Decimal(16, 1), TTrfQty)> 0 and TMultiIr<>''
	  	if @row>0 
			begin
	   		select @errmsg= 'Multi I/R Option Should Be Blank When Trf Qty Is Greater Than 0'
			goto ErrHndlr
	  	end

	  	select @row= count('x') from inserted where @wEntTyp= 'TxnBag' and TPrtKey=@wCurrPRtn and
			@wTcTyp Not IN ('BS') and TMultiIr <> '' 
	  	if @row>0 
			begin
	   		select @errmsg= 'Multi I/R Should Be Blank'
			goto ErrHndlr
	  	end

	  	select @row= count('x') from inserted ins where @wEntTyp= 'TxnBag' and TPrtKey=@wCurrPRtn and
			@wTcTyp IN ('BS') and 
			not exists (select 'x' from deleted del where ins.TIdNo=del.TIdNo and ins.TCoCd= del.TCoCd and ins.TTc= del.TTc 
				and ins.TYy= del.TYy and ins.TChr= del.TChr and ins.TNo= del.TNo and ins.TMultiIr= del.TMultiIr and ins.TPRtKey=del.TPRtKey)
			and exists(Select 'x' from Txnd where TdTIdNo=TIdNo and TdCoCd= TCoCd and TdTc= TTc 
				and TdYy= TYy and TdChr= TChr and TdNo= TNo and TdPrtKey=TPrtKEy) 

	  	if @row>0 
			begin
	   		select @errmsg= 'Cannot Change Multi I/R Option When Records Exist in Txnd'
			goto ErrHndlr
	  	end
			/***** Manali 3.8.0 - Multiple Issue/Receipt In BS  ****/


		/* Current Partition */
	  	select @row= count('x') from inserted where @wEntTyp in('TxnBag', 'TxnRm') and @wTcTyp 
			in ('DT', 'MLT', 'PR', 'MV', 'REC', 'CNV') and TPrtKey=@wCurrPrtn and (TFrRmLoc = '' or TToRmLoc = '') 
	  	if @row>0 
		begin
	   		select @errmsg= 'Default From/ To RM Location Cannot Be Blank'
			goto ErrHndlr
	  	end

		/* Current Partition */
	  	select @row= count('x') from inserted where @wEntTyp= 'TxnRm' and TPrtKey=@wCurrPrtn and TLsLoc<> '' and 
    			not exists (select 'x' from Loc where LocCoCd= TCoCd and LocTyp<> 'P' and LocCd= TLsLoc)
	  	if @row>0 
  		begin
	   		select @errmsg= 'Default Loss Loc Not Defined'
   			goto ErrHndlr
	  	end
  	
		/* Current Partition */
		select @row= count('x') from inserted where @wTcTyp= 'Rec' and TPrtKey=@wCurrPrtn and
			not exists (select 'x' from Loc where 
					LocCoCd= TCoCd and LocTyp= 'L' and LocCd= TLsLoc)
	  	if @row>0 
  		begin
   			select @errmsg= 'Default Loss Loc Not Defined'
	   		goto ErrHndlr
  		end
  	
		/* Current Partition */
		select @row= count('x') from inserted where @wEntTyp<> 'TxnRm' and TPrtKey=@wCurrPrtn
			and @wTcTyp<> 'REC' and TLsLoc<> '' 
	  	if @row>0 
  		begin
   			select @errmsg= 'Loss Loc Should Not Be Entered'
	   		goto ErrHndlr
  		end
  	
		/* Current Partition */
		select @row= count('x') from inserted where TPrtKEy=@wCurrPrtn and TLsToDt< TLsFrDt
	  	if @row>0 
  		begin
   			select @errmsg= 'To Loss Date Should Be Greater Than From Loss Date'
	   		goto ErrHndlr
  		end

		/*** In the 'REC' Entry, the TLsFrDt and TLsToDt range should be unique ***/
		if @wTcTyp='REC' 
		Begin
			/* Current Partition - Entry Not used */
			select @row= count('x') from inserted ins where TPrtKey=@wCurrPrtn and
				(Select count('x') from Txn T join Param on PTyp='TC' and PMCd=T.TTc 
					where PValue= 'REC' and T.TCoCd= ins.TCoCd and T.TLsLoc= ins.TLsLoc 
					and (        (T.TLsFrDt<= ins.TLsFrDt 
						      and T.TLsToDt>= ins.TLsFrDt)
					     or      (T.TLsFrDt<= ins.TLsToDt 
						      and T.TLsToDt>= ins.TLsToDt)
     					     or      (ins.TLsFrDt<= T.TLsFrDt 
						      and ins.TLsToDt>= T.TLsToDt)
	     				    )	
	   			) >1 
		  	if @row>0 
  			begin
	   			select @errmsg= 'Date Range Already Defined'
	   			goto ErrHndlr
		  	end
		End 

/*  ######################  Manoj #### Ver: 2.0.6 #### Date: 21/01/2004  ######################  */
		/* Current Partition */
	  	select @row= count('x') from Inserted Ins where Ins.TPrtKey=@wCurrPrtn and Ins.TSuppCd<> '' and TTc <> 'SA' and TTc<>'MV'
	And
			Ins.TSuppCd <> IsNull((Select Del.TSuppCd From Deleted Del Where
				Ins.TIdNo=Del.TIdNo and Ins.TCoCd = Del.TCoCd And Ins.TTc = Del.TTc And Ins.TYy = Del.TYy And
				Ins.TChr = Del.TChr And Ins.TNo = Del.TNo and Ins.TPrtKey=Del.TPrtKey), '')
		 And Not Exists (select 'x' from CustMst where CmCtg= 'S' and CmCd= Ins.TSuppCd And CmValidYN = 'Y')
 	  	if @row>0
  		begin
	   		select @errmsg= 'Supplier Code Not Defined / Invalid'
   			goto ErrHndlr
	  	end
/* (Bef 2.06)
	  	select @row= count('x') from inserted where TSuppCd<> '' and TTc <> 'SA' and
    			not exists (select 'x' from CustMst where CmCtg= 'S' and CmCd= TSuppCd)
	  	if @row>0
  		begin
	   		select @errmsg= 'Supplier Code Not Defined'
   			goto ErrHndlr
	  	end
*/
		/* Current Partition */
	  	select @row= count('x') from Inserted Ins where Ins.TPrtKey=@wCurrPrtn and Ins.TSuppCd<> '' and TTc= 'SA' and TTc='MV' And 
			Ins.TSuppCd <> IsNull((Select Del.TSuppCd From Deleted Del Where
				Ins.TIdNo=Del.TIdNo and Ins.TCoCd = Del.TCoCd And Ins.TTc = Del.TTc And Ins.TYy = Del.TYy And
				Ins.TChr = Del.TChr And Ins.TNo = Del.TNo And Ins.TPrtKey=Del.TPrtKey), '')
		 And Not Exists (select 'x' from CustMst where CmCtg= 'C' and CmCd= Ins.TSuppCd And CmValidYN = 'Y')
 	  	if @row>0
  		begin
	   		select @errmsg='Customer Code Not Defined / Invalid'
   			goto ErrHndlr
	  	end

	/* (Bef 2.06)
	  		select @row= count('x') from inserted where TSuppCd<> '' and TTc = 'SA' and
    				not exists (select 'x' from CustMst where CmCtg= 'C' and CmCd= TSuppCd)
	  		if @row>0
  			begin
	   			select @errmsg= 'Customer Code Not Defined'
   				goto ErrHndlr
	  		end
	*/

	/*	Date: 22/12/2003																			  */
		/******  TLockYN =  valid  PMCd ('YN')  ******/
		/* Current Partition */
		select @row= count('x') from inserted where @wEntTyp in ('TZ', 'TxnBag', 'TxnRm') And TPrtKey=@wCurrPrtn and
			Not Exists (select 'x' from Param where PTyp= 'YN' and PmCd= TLockYN)
		if @row> 0
		Begin
			select @errmsg= 'Custom Purchase Lock (Y/N) Not Defined' 
			goto ErrHndlr
		End

	/*		select @row= count('x') from inserted where @wEntTyp in ('TxnBag', 'TxnRm') And TLockYN<> 'N' 
		if @row> 0
		Begin
			select @errmsg= 'Txn Lock (Y/N) Should be 'N'' 
			goto ErrHndlr
		End
*/
		/* Current Partition */
		select @row= count('x') from inserted where @wEntTyp Not in ('TZ', 'TxnBag', 'TxnRm') and TPrtKey=@wCurrPrtn 
			And TLockYN<> ''
		if @row> 0
		Begin
			select @errmsg= 'Txn Lock (Y/N) Should be Blank'
			goto ErrHndlr
		End

/*  #########################################  Manoj  #########################################  */


/*	  	select @row= count('x') from inserted where @wEntTyp in ('TxnRm', 'Tz') and TSuppCd<> '' and 
    			not exists (select 'x' from CustMst where CmCtg= 'S' and CmCd= TSuppCd)
	  	if @row>0 
  		begin
	   		select @errmsg= 'Supplier Code Not Defined'
   			goto ErrHndlr
	  	end
*/
		/* Current Partition */
	  	select @row= count('x') from inserted where TPRtKey=@wCurrPrtn and TSuppCd= '' and TTc = 'SA' 
	  	if @row>0 
  		begin
	   		select @errmsg= 'Customer Code Should Be Entered'
   			goto ErrHndlr
	  	end

		/* Current Partition */
	  	select @row= count('x') from inserted where @wTcTyp in ('TZ', 'PR') and TPrtKey=@wCurrPrtn and TSuppCd= '' 
  		if @row>0 
	  	begin
  			select @errmsg= 'Supplier Code Should Be Entered'
	   		goto ErrHndlr
  		end

		/* Current Partition */
	  	select @row= count('x') from inserted where @wTcTyp not in ('MV','TZ', 'PR') and TPrtKey=@wCurrPrtn and TSuppCd<> '' 
  		if @row>0 
	  	begin
   			select @errmsg= 'Supplier Code Should Not Be Entered'
	   		goto ErrHndlr
  		end

		/* Current Partition */
		select @row= count('x') from inserted where @wEntTyp = 'Tz' and TPrtKey=@wCurrPrtn
			and TFrRmLoc<> '' and TToRmLoc<> '' and TFrBLoc<> '' and TLsLoc<> '' and TLsFrDt not in('01/01/1980') 
			and TLsToDt not in('01/01/1980') and TGldInLs<> 0 and TGldRec<> 0 and TDustWt<> 0 
	  	if @row> 0 
  		begin
   			select @errmsg= 'FrRmLoc/ ToRmLoc/ FrBagLoc/ Loss Loc/ FrLossDt/ ToLossDt/ Gld In Loss/ Gld Recovered/ Dust Wt Should Not Be Entered For Custom Purchase'
	   		goto ErrHndlr
  		end

		/* Old Code
	  			select @row= count('x') from inserted ins where TBoeNo<> '' and 
					(Select count('x') from Txn T where T.TBoeNo= ins.TBoeNo) >1 
	  			if @row>0 
  				begin
	   				select @errmsg= 'This Bill Of Entry Number Already Exists'
   					goto ErrHndlr
	  			end
		Old Code */

		/* BOE Number Unique Check on Complete Database */
	  	select @row= count('x') from inserted ins where TPrtKey=@wCurrPrtn and TBoeNo<> '' and 
			(Select count('x') from Txn T 
				join Param P1 on P1.PTyp= 'TC' and P1.PMCd= T.TTc and P1.PSCd= '' 
 				join Param P2 on P2.PTyp= 'TCTYP' and P2.PMCd= P1.PValue and P2.PSCd= '' 
				where P2.PValue1= @wEntTyp and T.TCoCd= ins.TCoCd and T.TBoeNo= ins.TBoeNo) >1 
	  	if @row>0 
  		begin
	   		select @errmsg= 'This Bill Of Entry Number Already Exists'
   			goto ErrHndlr
	  	end

		/* ****** Sachin 2.13.0 - Multi-Currency ****** */
		/* Current Partition */
	  	select @row= count('x') from Inserted Ins where TPrtKey=@wCurrPrtn and TTc = 'PR'
		 	And Not Exists 
			(Select 'x' from CustMst where CmCtg= 'S' and CmCd= Ins.TSuppCd And CmCurCd=Ins.TCurrCd )
 	  	if @row>0
  		begin
	   		select @errmsg= 'Supplier Currency Code Not Defined / Invalid'
   			goto ErrHndlr
	  	end

		/* Current Partition */
	  	select @row= count('x') from Inserted Ins where TPrtKey=@wCurrPrtn and TTc = 'SA'
		 	And Not Exists 
			(Select 'x' from CustMst where CmCtg= 'C' and CmCd= Ins.TSuppCd And CmCurCd=Ins.TCurrCd )
 	  	if @row>0
  		begin
	   		select @errmsg= 'Customer Currency Code Not Defined / Invalid'
   			goto ErrHndlr
	  	end

		/* Zubin 213 */
		/* Current Partition */
		select @row= count('x') from inserted where @wTcTyp in ('PR', 'TZ') AND TPrtKey=@wCurrPrtn and
			Convert(Decimal(16, 4), TCnvRt)<= 0
		if @row> 0 
		begin
			select @errmsg= 'Currency Conversion Rate Should Be > 0'
			goto ErrHndlr
		end

		/* Zubin 213 */
		/* Current Partition */
		select @row= count('x') from inserted where @wTcTyp in ('PR', 'TZ') AND TPRtKey=@wCurrPrtn and
			not exists (select 'x' from Param where PTyp= 'MULDIV' and PMCd= TMulDiv and PSCd= '')
		if @row> 0 
		begin
			select @errmsg= 'Currency Multiplication Factor Not Defined'
			goto ErrHndlr
		end
		/* ****** Sachin 2.13.0 - Multi-Currency ****** */
	End

	if @mode in ('C')
	Begin
		/* Sachin 2.12 */ 
		/*** Cannot change the Tran Date when Txnd Recs exists ***/
		/* Current Partition */
/*		select @row= count('x') from inserted ins where ins.TPRtKey=@wCurrPrtn and
			not exists (select 'x' from deleted del where ins.TIdNo=del.TIdNo and ins.TCoCd= del.TCoCd and ins.TTc= del.TTc 
				and ins.TYy= del.TYy and ins.TChr= del.TChr and ins.TNo= del.TNo and ins.TDt= del.TDt and ins.TPRtKey=del.TPRtKey)
			and exists(Select 'x' from Txnd where TdTIdNo=TIdNo and TdCoCd= TCoCd and TdTc= TTc 
				and TdYy= TYy and TdChr= TChr and TdNo= TNo and TdPrtKey=TPrtKEy) 
		if @row> 0 
		begin
			select @errmsg= 'Txn Date Cannot Be Changed When Txnd Rec exists'
			goto ErrHndlr
		end
*/
		/* Sachin 2.12 */ 

		/*** Cannot change the Primary Key ***/
		/* Current Partition */
		select @Row=Count('x') from deleted del where del.TPrtKEy=@wCurrPrtn and
			not exists (Select 'x' from inserted ins where ins.TIdNo=del.TIdNo and ins.TCoCd= del.TCoCd 
			and ins.TTc= del.TTc and ins.TYy= del.TYy and ins.TChr= del.TChr 
			and ins.TNo= del.TNo and ins.TPrtKey=del.TPrtKey)
  		If @Row > 0
		begin    
			Select @ErrMsg='Cannot Change Voucher No.'
			GoTo ErrHndlr
 		end

		/*** Cannot change Default Bag Loc, RM Fr & To Locs when Txnd Rec exists ***/
		/* Current Partition */
		select @Row=Count('x') from deleted del where del.TPrtKey=@wCurrPrtn and
			not exists (Select 'x' from inserted ins where ins.TIdNo=del.TIdNo and ins.TCoCd= del.TCoCd and ins.TTc= del.TTc 
			and ins.TYy= del.TYy and ins.TChr= del.TChr and ins.TNo= del.TNo and ins.TPrtKey=del.TPRtKey and ins.TFrBLoc= del.TFrBLoc 
			and ins.TFrRmLoc= del.TFrRmLoc and ins.TToRmLoc= del.TToRmLoc 
			and ins.TTrfQty= del.TTrfQty) and 
			exists (Select 'x' from Txnd where TdTIdNo=del.TIdNo and TdCoCd= del.TCoCd and TdTc= del.TTc 
			and TdYy= del.TYy and TdChr= del.TChr and TdNo= del.TNo and TdPrtKey=del.TPrtKey)
  		If @Row> 0
		begin
			Select @ErrMsg= 'Cannot Change Default Bag Loc/ RM Fr-To Loc/ Cross Transfer Qty When Txnd Record Exists'
			GoTo ErrHndlr
 		end

		/******  Cannot change TImpTyp when Child Record Exists ******/ 
/*		select @Row= Count('x') from deleted del where @wEntTyp in ('Tz') and 
			not exists (Select 'x' from inserted ins where ins.TCoCd= del.TCoCd and ins.TTc= del.TTc and ins.TYy= del.TYy 
				    and ins.TChr= del.TChr and ins.TNo= del.TNo and ins.TImpTyp= del.TImpTyp) 
			and exists (Select 'x' from Txndz where TzCoCd= del.TCoCd and TzTc= del.TTc and TzYy= del.TYy 
				    and TzChr= del.TChr and TzNo= del.TNo) 
  		If @Row > 0
		begin    
			Select @ErrMsg= 'Cannot Change Import Type When Child Record Exists'
			GoTo ErrHndlr
 		end
*/
		/******  Cannot change TUseBefDt when Custom Match Records Exist ******/ 
		/* Current Partition */
		select @Row=Count('x') from deleted del where @wEntTyp in ('Tz') and TPRtKey=@wCurrPrtn and
			not exists (Select 'x' from inserted ins where ins.TIdNo=del.TIdNo and ins.TCoCd= del.TCoCd and ins.TTc= del.TTc and ins.TYy= del.TYy 
				    and ins.TChr= del.TChr and ins.TNo= del.TNo and ins.TPrtKey=del.TPrtKEy and ins.TUseBefDt= del.TUseBefDt) 
			and exists (Select 'x' from Txndz where TzTIdNo=del.TIdNo and TzCoCd= del.TCoCd and TzTc= del.TTc and TzYy= del.TYy 
				    and TzChr= del.TChr and TzNo= del.TNo and TzPrtKey=del.TPrtKey and Round(TzRmDespWt, 3)> 0 )
  		If @Row > 0
		begin    
			Select @ErrMsg= 'Cannot Change Use Before Date When Custom Match Record Exists '
			GoTo ErrHndlr
 		end

/* **** Zubin 212 **** */ 
		/* Current Partition */
	  	select @row= count('x') from inserted ins 
			where @wEntTyp= 'TxnRm' and @wTcTyp IN ('PR', 'MV') and TPrtKey=@wCurrPrtn and
				Not Exists (Select 'x' from deleted del where ins.TIdNo=del.TIdNo and ins.TCoCd= del.TCoCd and 
					ins.TTc= del.TTc and ins.TYy= del.TYy and ins.TChr= del.TChr and 
					ins.TNo= del.TNo and ins.TPrtKey=del.TPrtKey and ins.TLotNo= del.TLotNo) 
				and Exists (Select 'x' from Txnd where TdTIdNo=TIdNo and TdCoCd= TCoCd and TdTc= TTc and TdYy= TYy and 
					TdChr= TChr and TdNo= TNo and TdPrtKey=TPrtKey) 
	  	if @row>0 
		begin
	   		select @errmsg= 'Cannot Change Lot No. As Child Record Exists'
			goto ErrHndlr
	  	end
/* **** Zubin 212 **** */ 


/* **** Sachin 2.13.0 **** */ 
		/* Current Partition */
	  	select @row= count('x') from inserted ins 
			where @wEntTyp= 'TxnRm' and @wTcTyp IN ('PR', 'MV') and TPrtKey=@wCurrPrtn and
				Not Exists (Select 'x' from deleted del where ins.TIdNo=del.TIdNo and ins.TCoCd= del.TCoCd and 
					ins.TTc= del.TTc and ins.TYy= del.TYy and ins.TChr= del.TChr and 
					ins.TNo= del.TNo and ins.TPrtKey=del.TPrtKey and convert(decimal (16, 4), ins.TCnvRt)= convert(decimal (16, 4), del.TCnvRt)) 
				and Exists (Select 'x' from Txnd where TdTIdNo=TIdNo and TdCoCd= TCoCd and TdTc= TTc and TdYy= TYy and 
					TdChr= TChr and TdNo= TNo and TdPrtKey=TPrtKey) 
	  	if @row>0 
		begin
	   		select @errmsg= 'Cannot Change Conversion Factor As Child Record Exists'
			goto ErrHndlr
	  	end
/* **** Sachin 2.13.0 **** */ 
	End

	-- 4.1.3.0
	if @mode='C' And Update(TDt)
	Begin
		if @wTcTyp ='PR'
			Update Txnd Set TdDt=TDt From Inserted Where TCoCd=TdCoCd And TTc=TdTc And TYy=TdYy And TChr=TdChr And
					TNo=TdNo And TdDt<>TDt
	End

	if @mode in ('D')
	Begin
		/*** Cannot Delete Transaction If Child Record Exists (Txnd) ***/
		/* Current Partition */
		select @row=count('x') from deleted join Txnd on TIdNo=TdTIdNo and TCoCd= TdCoCd and TTc= TdTc 
			and TYy= TdYy and TChr= TdChr and TNo= TdNo and TPrtKey=@wCurrPrtn 
		if @row>0 
  		begin
   			select @errmsg='Cannot Delete Transaction As Child Record Exists'
   			goto ErrHndlr
  		end

		/*** Cannot Delete Transaction If Child Record Exists (TxndZ) ***/
		/* Current Partition */
		select @row=count('x') from deleted join TxndZ on TIdNo=TzTIdNo and TCoCd= TzCoCd and TTc= TzTc and TYy= TzYy and 
			TChr= TzChr and TNo= TzNo and TPrtKey=@wCurrPrtn
		if @row>0 
  		begin
   			select @errmsg='Cannot Delete Purchase Entry As Child Record Exists'
   			goto ErrHndlr
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


Alter TRIGGER [dbo].[DsgAna_INUPDL] ON [dbo].[DsgAna] 
	FOR INSERT,UPDATE,DELETE NOT FOR REPLICATION
AS 
BEGIN 
	/* *** Trigger Version Emr 3.02 [Speed Optimization - Purge Partition] *** */
	DECLARE @ROW INT 
	DECLARE @mode char(1) 
	DECLARE @errmsg varchar(255) 
	DECLARE @ReplAgt BIT  
	Declare @wCurrPrtn VarChar(1)
	
	Select @wCurrPrtn = 'C'

	IF @@ROWCOUNT=0 RETURN 
 
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'DsgAna' ) 
		RETURN 
 
	select @mode=case (select count('x') from inserted) when 0 then 'D' 
	else 
		case (select count('x') from deleted) when 0 then 'A' else 'C' end 
	end 
 
	 /* **** Checking Whether a Normal Transaction or Replicated One **** */  
	 Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))  

	 If @ReplAgt = 1         Return
      

	if @mode='C'  And @ReplAgt <> 1
 	Begin
	/******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */
		Select @row=count('x') From Inserted Where DaPrtKey<>@wCurrPrtn 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Edit Records of Previous Partition'
			goto ErrHndlr
		end
	end

	if @mode='A'  And @ReplAgt <> 1
 	Begin
		/******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */
		Select @row=count('x') From Inserted Where DaPrtKey <> @wCurrPrtn 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Insert Partition Key Other Than ''C'''
			goto ErrHndlr
		end
	End

 	if @mode = 'D'  And @ReplAgt <> 1
 	Begin
  		select @Row=Count('x') from deleted del where del.DaPrtKey<>@wCurrPrtn
  		If @Row > 0
	  	begin
			Select @ErrMsg='Cannot Delete Records From Previous Partition'
			GoTo ErrHndlr
		end
	End
	
	If  (@ReplAgt<>1)
	begin
		/******Cannot Insert/Update/Delete if Crm database /Emr database with sdMst='N' ****/
		if @mode in ('A','C','D')
		Begin 
		If   exists ( select 'x' From SysDb where SysId='EMR' and SdMst='N' )  or  
		exists ( select 'x' From SysDb where SysId='CRM' and SysDb ='WEB' ) 
			Begin 
				select @errmsg= 'Cannot Insert /Update /Delete ' 
				goto ErrHndlr 
			end 
		End 
	end	
 	/******** Key Cannot Be Changed ********/
 	if @mode = 'C'
 	Begin
  		select @Row= Count('x') from deleted del where
    		Not Exists (Select 'x' from inserted ins where ins.DaIdNo=del.DaIdNo and ins.DaTcTyp= del.DaTcTyp 
					and ins.DaCd= del.DaCd And ins.DaSz= del.DaSz and ins.DaAnaSr= del.DaAnaSr)
  		If @Row> 0
	  	begin
			Select @ErrMsg= 'Cannot Change DsgAna Key'
			GoTo ErrHndlr
		end
	End

	/****** DaTcTyp = 'DM' ******/
	/* Partition Check Not Required */
/*	select @row= count('x') from inserted where DaTcTyp<> 'DM'
	if @row> 0
	begin
		select @errmsg= 'Analysis Fields Canbe Entered Only In Design Master'
		goto ErrHndlr
	end
*/
	/****** DaTcTyp, DaCd, DaSz combination should exist in the DsgMst ******/
	/* Check On Current Partition */
	select @row= count('x') from inserted Ins where 
		Not Exists (Select 'x' From Deleted Del Where Ins.DaDmIdNo=Del.DaDmIdNo and Ins.DaTcTyp=Del.DaTcTyp 
						and Ins.DaCd=Del.DaCd and Ins.DaSz=Del.DaSz and Ins.DaPrtKey<>Del.DaPrtKey) and
		not exists (select 'x' from DsgMst where
				DmIdNo=INs.DaDmIdNo and DmTcTyp= Ins.DaTcTyp and DmCd= Ins.DaCd and DmSz= Ins.DaSz and DmPrtKey=Ins.DaPrtKey)
	if @row> 0
	begin
		select @errmsg= 'This Type-Design-Size Combination Not Defined In Design Master'
		goto ErrHndlr
	end

	/******  DaAnaSr  =  PMCd ('DAANAFLD')  ******/
	/* Partition Check Not Required */
	select @row= count('x') from inserted where
		not exists (select 'x' from Param where PTyp= 'DAANAFLD' and PMCd= DaAnaSr)
	if @row> 0
	begin
		select @errmsg= 'Design Analysis Field Not Defined'
		goto ErrHndlr
	end

	/******  DaAnaCd  =  PMCd ('DAANACD')  ******/
	/* Partition Check Not Required */
	select @row= count('x') from inserted where
		(	exists (Select 'x' from Param where PTyp= 'DAANAFLD' and PMCd= DaAnaSr
			and PValue= 'N')
			And DaAnaCd<> ''
		)
		OR
		(	DaAnaCd<> ''
			And exists (Select 'x' from Param where PTyp= 'DAANAFLD' and PMCd= DaAnaSr
				and PValue= 'Y')
			And not exists (select 'x' from Param where PTyp= 'DAANACD' and PMCd= DaAnaSr and PSCd= DaAnaCd)
		)
		OR
		(	DaAnaCd= ''
			And exists (Select 'x' from Param where PTyp= 'DAANAFLD' and PMCd= DaAnaSr
				and PValue= 'Y' and PValue1= 'Y')
		)
		if @row> 0
		begin
			select @errmsg= 'Design Analysis Code Not Defined'
			goto ErrHndlr
		end

	/******  DaAnaDesc  ******/
	/* Partition Check Not Required */
	select @row= count('x') from inserted where
		(	exists (Select 'x' from Param where PTyp= 'DAANAFLD' and PMCd= DaAnaSr
			and PValue= 'Y')
			And DaAnaDesc<> ''
		)
		OR
		(	DaAnaDesc= ''
			And exists (Select 'x' from Param where PTyp= 'DAANAFLD' and PMCd= DaAnaSr
				and PValue= 'N' and PValue1= 'Y')
		)
		if @row> 0
		begin
			select @errmsg= 'Design Analysis Description Not Defined'
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

If not exists (select 'x' from sys.indexes where  upper(name) = 'NC_ODBASEIDKEY' and upper(object_name(object_id)) ='ORDDSG') 
Begin
	CREATE NONCLUSTERED INDEX [NC_OdBaseIdKey] ON [dbo].[OrdDsg]
		(
			[OdPrtKey] ASC,
			[OdBaseIdKey] ASC
		) include 	([odcocd] )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
End


GOGO

If not exists (select 'x' from sys.indexes where  upper(name) = 'NC_ODREFIDKEY' and upper(object_name(object_id)) ='ORDDSG') 
begin
	CREATE NONCLUSTERED INDEX [NC_OdRefIdKey] ON [dbo].[OrdDsg]
	(
		[OdPrtKey] ASC,
		[OdCoCd] ASC,
		[OdRefIdKey] ASC
	)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
End
GOGO


If exists (select 'x' from sys.indexes where  upper(name) = 'NC_INVFGD' and upper(object_name(object_id)) ='INVFGD') 
begin
	drop index nc_invfgd on invfgd
End

If exists (select 'x' from sys.objects where upper(name) = 'PK_INVFGD' and upper(object_name(parent_object_id)) ='INVFGD') 
Begin
	ALTER TABLE [dbo].[InvFgd] drop  CONSTRAINT [PK_InvFgd]  
end

GOGO

/****** Object:  Index [PK_InvFgd]    Script Date: 8/8/2015 11:45:20 PM ******/
ALTER TABLE [dbo].[InvFgd] ADD  CONSTRAINT [PK_InvFgd] PRIMARY KEY CLUSTERED 
(
		[IfIdNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]

GOGO


/****** Object:  Index [NC_InvFgd]    Script Date: 8/8/2015 11:43:12 PM ******/
CREATE NONCLUSTERED INDEX [NC_InvFgd] ON [dbo].[InvFgd]
(
	[IfCoCd] ASC,
	[IfTc] ASC,
	[IfYy] ASC,
	[IfChr] ASC,
	[IfNo] ASC,
	[IfSrNo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
GOGO

If exists (select 'x' from sys.indexes where  upper(name) = 'NC_INVFGD2' and upper(object_name(object_id)) ='INVFGD') 
begin
	drop index nc_invfgd2 on invfgd
End

/****** Object:  Index [NC_InvFgd2]    Script Date: 8/8/2015 11:43:32 PM ******/
CREATE NONCLUSTERED INDEX [NC_InvFgd2] ON [dbo].[InvFgd]
(
	[IfCoCd] ASC,
	[IfTc] ASC,
	[IfYy] ASC,
	[IfChr] ASC,
	[IfNo] ASC,
	[IfSr] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
GOGO


insert into trgdisa values ('param');
gogo
if not exists (select 'a' from param where ptyp = 'MTCHQTY') 
begin
	Insert into param
(PTyp     ,
PMCd     ,
PSCd     ,
PDesc                          ,
PDesc225                                                                                                                                                                                                                          ,
PValue                         ,
PNum                   ,
PValue1                                                                                                                  ,
PNum1                  ,
PValue2                        ,
ModUsr ,
ModDt                   ,
ModTime                                 ,
PValue3                        ,
PValidYn ,
PPrtKey)
values
('PTYP',
'MTCHQTY',
'',
'Custom Match Qty Check',
'',
'Y',
0,
'',
0,
'',
'MIC',
'1980-01-01 00:00:00',
0.00,
'',
'N',
'C')

Insert into param
(PTyp     ,
PMCd     ,
PSCd     ,
PDesc                          ,
PDesc225                                                                                                                                                                                                                          ,
PValue                         ,
PNum                   ,
PValue1                                                                                                                  ,
PNum1                  ,
PValue2                        ,
ModUsr ,
ModDt                   ,
ModTime                                 ,
PValue3                        ,
PValidYn ,
PPrtKey)
values
('MTCHQTY',
'MTCHQTY',
'',
'Custom Match Qty Check',
'',
'Y',
0,
'',
0,
'',
'MIC',
'1980-01-01 00:00:00',
0.00,
'',
'N',
'C')

end

Update param set pdesc225 = '5.0.2' where ptyp = 'VER' and pmcd = 'DATVER';

Delete from trgdisa;

gogo







