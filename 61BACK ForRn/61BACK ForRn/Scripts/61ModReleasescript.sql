
ALTER TRIGGER [dbo].[RmMst_INUPDL] ON [dbo].[RmMst] 
     FOR INSERT,UPDATE,DELETE
	NOT FOR REPLICATION
AS
BEGIN
	/* *** Trigger Version Emr 3.10.0 *** */
	DECLARE @ROW INT
	DECLARE @mode char(1)
	DECLARE @errmsg varchar(255)
	DECLARE @SYS VARCHAR(5)
	Declare @wHCoCd Varchar(8) 
	Declare @wCurrPrtn varchar(1)
	DECLARE @ReplAgt BIT  

	IF @@ROWCOUNT= 0
		RETURN
	If Exists(Select 'x' From TrgDisa Where TdTblNm = 'RmMst' )
		RETURN

	Set @SYS= IsNull((Select PMCd from Param where PTyp= 'SYS' and PMCd= 'XEMR' and PSCd= ''), '') 

	select @mode= case (select count('x') from inserted) when 0 then 'D'
			 else
				case (select count('x') from deleted) when 0 then 'A' else 'C' end
			 end 

	Select @wHCoCd= 'ZZZ'
	Select @wCurrPrtn= 'C'

	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))

	/******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */
 	if @mode in ('A','C') And @ReplAgt <> 1
 	Begin		
		Select @row=count('x') From Inserted Where RmPrtKey<>@wCurrPrtn
		if @row> 0 
		begin
			select @errmsg= 'Cannot Edit Records of Previous Partition'
			goto ErrHndlr
		end

		Select @row=count('x') From Inserted Where RmPrtKey <> @wCurrPrtn 
		if @row> 0 
		begin
			select @errmsg= 'Cannot Insert Partition Key Other Than ''C'''
			goto ErrHndlr
		end
	End

	if @mode = 'D' And @ReplAgt <> 1 
	Begin
		Select @row= count('x') From deleted del Where del.RmPrtKey<>@wCurrPrtn
		If @row> 0 
		Begin
			select @errmsg= 'Cannot Delete Records From Previous Partition'
			goto ErrHndlr
		End
	End
	/******* Sachin 3.02.0 [30/11/07] - Purge Partition (Speed Optimization) */


 	If @mode = 'C'
 	Begin
		/******** Primary Key (RmCd) along with RmCtg and RmSCtg Cannot Be Changed ********/ 
		/* Check on Current Db */
		select @Row= Count('x') from deleted del where del.RmPrtKey=@wCurrPrtn and
			not exists (Select 'x' from inserted ins where ins.RmCd= del.RmCd and 
				    ins.RmCtg= del.RmCtg and ins.RmSCtg= del.RmSCtg)
		If @Row > 0
		begin 
			Select @ErrMsg= 'Cannot Change Rm Code/ Category/ SubCategory'
			GoTo ErrHndlr
		end

		/******** Cannot Change Customs Code When InvRm Record Exists ********/ 
		/* Check on Current Db */
		select @Row= Count('x') from deleted del where RmZ= 'N' and RmPrtKey=@wCurrPrtn and 
			not exists (Select 'x' from inserted ins where ins.RmCd= del.RmCd and ins.RmPrtKey=del.RmPrtKey and ins.RmZCd= del.RmZCd)
			and exists(Select 'x' from InvRm where IrRmCd= del.RmCd)
		If @Row > 0
		begin 
			Select @ErrMsg= 'Cannot Change Customs Code When InvRm Record Exists'
			GoTo ErrHndlr
		end

		/******** Cannot Change RmQw When RmRt Record Exists ********/ 
		/* Check on Current Db */
		Select @Row= Count('x') from deleted del where del.RmPrtKey=@wCurrPrtn and 
			not exists (Select 'x' from inserted ins where ins.RmCd=del.RmCd and ins.RmQw=del.RmQw and ins.RmQwCst=del.RmQwCst and ins.RmPrtKey=@wCurrPrtn)
			and exists (Select 'x' from RmRt where RrTcTyp= 'RM' and RrCd= del.RmCd and RrPrtKey=RmPrtKey) 
		If @Row > 0
		begin 
			Select @ErrMsg= 'Cannot Change Sale/Cost Q/W As Rm Rate Record Exists'
			GoTo ErrHndlr
		end	

		/******** Cannot Change RmQw When OrdRm Record Exists ********/ 
		/* Check on Current Db */
		select @Row= Count('x') from deleted del where del.RmPrtKey=@wCurrPrtn and
			not exists (Select 'x' from inserted ins where ins.RmCd= del.RmCd and ins.RmQw= del.RmQw and ins.RmQwCst=del.RmQwCst and ins.RmPrtKey=del.RmPrtKey )
			and exists(Select 'x' from OrdRm where OrRmCd= del.RmCd and OrPrtKey=RmPrtKey) 
		If @Row > 0
		begin 
			Select @ErrMsg= 'Cannot Change Sale/Cost Q/W As Order Rm Record Exists'
			GoTo ErrHndlr
		end

		/******** Manali 3.10.0 - Cannot Change RmQw When MultiPrcQtRm Record Exists ********/ 
		/* Check on Current Db */
		select @Row= Count('x') from deleted del where del.RmPrtKey=@wCurrPrtn and
			not exists (Select 'x' from inserted ins where ins.RmCd= del.RmCd and ins.RmQw= del.RmQw and ins.RmQwCst=del.RmQwCst and ins.RmPrtKey=del.RmPrtKey )
			and exists(Select 'x' from MultiPrcQtRm where MqrRmCd= del.RmCd and MqrPrtKey=RmPrtKey) 
		If @Row > 0
		begin 
			Select @ErrMsg= 'Cannot Change Sale/Cost Q/W As MultiPrcQtRm Record Exists'
			GoTo ErrHndlr
		end

		/******** Cannot Change RmQw When Txnd Record Exists ********/ 
		/* Check on Current Db */
		select @Row= Count('x') from deleted del where del.RmPrtKey=@wCurrPrtn and 
			not exists (Select 'x' from inserted ins where ins.RmCd= del.RmCd and ins.RmPrtKey=del.RmPrtKey and ins.RmQw= del.RmQw)
			and exists(Select 'x' from Txnd where TdRmCd= del.RmCd) 
		If @Row > 0
		begin 
			Select @ErrMsg= 'Cannot Change Q/W As Txnd Record Exists'
			GoTo ErrHndlr
		end

		/* ****** Sachin 2.12 - 06/12/05 - SJM ****** */ 
			/****** RmPurityWt cannot be changed for Silver if HSlvModYN is Set to 'N' in Head ******/
			/* Check on Current Db */
			select @row= count('x') from inserted ins where ins.RmCtg='S' and ins.RmPrtKey=@wCurrPrtn and
				Not Exists (Select 'x' from deleted del where ins.RmCd= del.RmCd and ins.RmPurityWt= del.RmPurityWt and ins.RmPrtKey=del.RmPrtKey) and 
				Exists (Select 'x' from Head where HCoCd=@wHCoCd And HSlvModYN='N') 				
			if @row> 0 
			begin
				select @errmsg= 'Cannot Change Prd Purity if Silver Module option is set to ''N'' in Head' 
				goto ErrHndlr
			end 
		/* ****** Sachin 2.12 - 06/12/05 - SJM ****** */ 
		
		/* ***** Manali 3.03 - 17/06/08 - Palladium Module */		
		/****** RmPurityWt cannot be changed for Palladium if HPdModYN is Set to 'N' in Head ******/
		/* Check on Current Db */
		select @row= count('x') from inserted ins where ins.RmCtg='L' and ins.RmPrtKey=@wCurrPrtn and
			Not Exists (Select 'x' from deleted del where ins.RmCd= del.RmCd and ins.RmPurityWt= del.RmPurityWt and ins.RmPrtKey=del.RmPrtKey) and 
			Exists (Select 'x' from Head where HCoCd=@wHCoCd And HPdModYN='N') 				
		if @row> 0 
		begin
			select @errmsg= 'Cannot Change Prd Purity if Palladium Module option is set to ''N'' in Head' 
			goto ErrHndlr
		end 
		/* ***** Manali 3.03 - 17/06/08 - Palladium Module */		
	End

 	if @mode in ('A','C')
 	Begin
		/****** RmId = PSCD('RMID') ******/
		/* Current Partition  */
		select @row= count('x') from Inserted where RmId <> '' And RmZ= 'N' And RmPrtKey=@wCurrPrtn and
			not exists (select 'x' from Param where PTyp= 'RMID' and PMCd= RmCtg And PSCd = RmId)
		if @row> 0 
		begin
			select @errmsg= 'Raw Material Id  Not Defined'
			goto ErrHndlr
		end
		
		/****** RmId = PSCD('RMID') ******/
		/* Current Partition  */
		select @row= count('x') from Inserted where RmId <> '' And RmZ= 'Y'  And RmPrtKey=@wCurrPrtn
		if @row> 0 
		begin
			select @errmsg= 'Raw Material Id Should be Blank'
			goto ErrHndlr
		end
	
		/****** RmSCtg= PSCd ('RMSCTG', RmCtg) ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmPrtKey=@wCurrPrtn and not exists (select 'x' from Param where 
			PTyp= 'RMSCTG' and PMCd= RmCtg and PSCd= RmSCtg)
		if @row> 0 
		begin
			select @errmsg= 'Raw Material Sub Category Not Defined'
			goto ErrHndlr
		end
	
		/****** RmSTWGrp= PMCd ('STWGRP') when RmCtg= 'C' or 'D' and RmZ= 'N' ******/
		/* Current Partition  */
		--6.1 new 'G', 'P', 'S', 'L', 'M', 'X' added
		select @row= count('x') from inserted where RmCtg in ('C', 'D','G', 'P', 'S', 'L', 'M', 'X') and RmZ= 'N' and RmPrtKey=@wCurrPrtn and
			not exists (select 'x' from Param where PTyp= 'STWGRP' and PMCd= RmSTWGrp)
		if @row> 0 
		begin
			select @errmsg= 'Size To Weight Group Not Defined'
			goto ErrHndlr
		end
	
		/****** RmSTWGrp= Blank when RmCtg in ('G', 'P', 'S', 'A', 'M', 'X') or RmZ= 'Y' ******/
		/* Current Partition  */
		/* *** Manali 3.03 - 17/06/08 - 'L'  added */
		--6.1 new 'G', 'P', 'S', 'L', 'M', 'X'
		select @row= count('x') from inserted where RmSTWGrp<> '' and RmPrtKey=@wCurrPrtn and 
			(RmCtg in ('A') or RmZ= 'Y')  			
		if @row> 0 
		begin
			select @errmsg= 'Size To Weight Group Should Be Blank'
			goto ErrHndlr
		end
	
		/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
		/*	Date: 09/12/2003																			 */
		/****** RmValidYn = PMCd ('YN') ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmPrtKey=@wCurrPrtn and
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= RmValidYN)
		if @row> 0
		begin
			select @errmsg= 'Valid (Y/N) Not Defined'
			goto ErrHndlr
		End
		/*  #########################################  Manoj  #########################################  */
	
		/****** RmParentCd= PMCd ('PARENTCD') when RmCtg= 'C' or 'D' and RmZ= 'N' ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmParentCd<> '' and RmCtg in ('C', 'D') and RmZ= 'N' and RmPrtKey=@wCurrPrtn and 
			Not exists (select 'x' from Param where PTyp= 'PARENTCD' and PMCd= RmParentCd)
		if @row> 0 
		begin
			select @errmsg= 'Parent Code Not Defined'
			goto ErrHndlr
		end
	
		/****** RmParentCd= Blank when RmCtg in ('G', 'P', 'S', 'A', 'M', 'X') or RmZ= 'Y' ******/
		/* Current Partition  */
		/* *** Manali 3.03 - 17/06/08 - 'L'  added */
		select @row= count('x') from inserted where RmParentCd<>'' and RmPrtKey=@wCurrPrtn and
			(RmCtg in ('G', 'P', 'S', 'L', 'A', 'M', 'X') or RmZ= 'Y')  			
		if @row> 0 
		begin
			select @errmsg= 'Parent Code Should Be Blank'
			goto ErrHndlr
		end
	
		/* Current Partition  */		
		select @row= count('x') from inserted where RmCd= '' and RmPrtKey=@wCurrPrtn
		if @row> 0 
		begin
			select @errmsg= 'Raw Material Code Cannot Be Blank'
			goto ErrHndlr
		end
	
		/****** RmSpGrav <> 0 ******/
		/* Current Partition  */
		/* *** Manali 3.03 - 17/06/08 - 'L' added */
		select @row= count('x') from inserted where RmSpGrav= 0 and RmPrtKey=@wCurrPrtn and
			RmCtg in ('G', 'P', 'S', 'L', 'A', 'M') 
		if @row> 0 
		begin
			select @errmsg= 'Specific Gravity Cannot Be Zero'
			goto ErrHndlr
		end
	
		/****** RmPurityRt <> 0 ******/
		/* Current Partition  */
		/* *** Manali 3.03 - 17/06/08 - 'L' added */
		select @row= count('x') from inserted where RmPurityRt= 0 and RmPrtKey=@wCurrPrtn and
			RmCtg in ('G', 'P', 'S', 'L','M') 
		if @row> 0 
		begin
			select @errmsg= 'Purity Cannot Be Zero'
			goto ErrHndlr
		end
	
		/****** RmPurityWt <> 0 for RmCtg= G,P ******/
		/* Current Partition  */
		/* *** Manali 3.03 - 17/06/08 - 'L' added */
		select @row= count('x') from inserted where RmPurityWt= 0 and RmPrtKey=@wCurrPrtn and
			RmCtg in ('G', 'P', 'S', 'L','M') 
		if @row> 0 
		begin
			select @errmsg= 'Production Purity Cannot Be Zero'
			goto ErrHndlr
		end
	
		/* **** Zubin 211 **** */ 
			/****** RmPurityWt Cannot be changed RmCtg= G,P,S ******/
			/* Current Partition  */
			/* *** Manali 3.03 - 17/06/08 - 'L' added */
		select @row= count('x') from inserted ins where ins.RmCtg In ('G', 'P', 'S', 'L','M') and ins.RmPrtKey=@wCurrPrtn and 
				Not Exists (Select 'x' from deleted del where ins.RmCd= del.RmCd and ins.RmPurityWt= del.RmPurityWt and ins.RmPrtKey=del.RmPrtKey) and 
				Exists (Select 'x' from Txnd where TdRmCd= RmCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Change Production Purity As RmCd Is Used In Txnd Table' 
				goto ErrHndlr
			end
		/* **** Zubin 211 **** */ 

		/****** RmBasePurityWt <> 0 for RmCtg= X ******/
		/* Current Partition  */
		select @row= count('x') from inserted where Round(RmBasePurityWt, 4)= 0 and RmZ= 'N' and RmPrtKey=@wCurrPrtn 
			and RmCtg= 'X' 
		if @row> 0 
		begin
			select @errmsg= 'Base Production Purity Cannot Be Zero For Accessories'
			goto ErrHndlr
		end
	
		/****** RmBasePurityWt = 0 for RmZ= 'Y' or RmCtg = G,P,S,D,C,A,M ******/
		/* Current Partition  */
		/* *** Manali 3.03 - 17/06/08 - 'L' added */
		select @row= count('x') from inserted where Round(RmBasePurityWt, 4)<> 0 and RmPrtKey=@wCurrPrtn and
			(RmZ= 'Y' or RmCtg in ('G', 'P', 'S', 'L', 'D', 'C', 'A', 'M')) 
		if @row> 0 
		begin
			select @errmsg= 'Base Production Purity Should Be Zero For RmCtg = G,P,S,L,D,C,A,M Or Custom Code'
			goto ErrHndlr
		end
	
		/****** RmCmCd = Blank or valid CmCd from CustMst (CmCtg= 'C') ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmCmCd<> '' and RmPrtKey=@wCurrPrtn
			and not exists (select 'x' from CustMst where 
			CmCtg= 'C' and CmCd= RmCmCd)
		if @row> 0 
		begin
			select @errmsg= 'Customer Code Not Defined'
			goto ErrHndlr
		end

			
		/****** RmUom = PMCd ('RMUOM') ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmPrtKey=@wCurrPrtn and
			not exists (select 'x' from Param where PTyp= 'RMUOM' and PMCd= RmUom)
		if @row>0 
		begin
			select @errmsg= 'Raw Material Unit Of Measurement Not Defined'
			goto ErrHndlr
		end
	
	/*		Manoj 2.10.0		*/
		/****** RmSCtg = PSCd ('RMCOL', RmCtg) ******/
	/*	select @row= count('x') from inserted where not exists (select 'x' from Param where 
				PTyp= 'RMCOL' and PMCd= RmCtg and PSCd= RmCol)
		if @row> 0 
		begin
			select @errmsg= 'Raw Material Colour Not Defined'
			goto ErrHndlr
		end
	*/
		/* Current Partition  */
		select @row= count('x') from inserted Ins where Ins.RmPrtKey=@wCurrPrtn and 
			Ins.RmCol <> IsNull((Select Del.RmCol From Deleted Del Where Ins.RmCd = Del.RmCd and Ins.RmPrtKey=Del.RmPrtKey), '') And 
			not exists (select 'x' from Param where PTyp= 'RMCOL' and PMCd= RmCtg and PSCd= RmCol And PValidYn = 'Y')
		if @row> 0 
		begin
			select @errmsg= 'Raw Material Colour Not Defined / Invalid.'
			goto ErrHndlr
		end
	/*		Manoj 2.10.0		*/
		
		/****** If RmCtg= 'G' or 'P' or 'S' then RmKt = PMCd ('KT')  	else RmKt= Blank or PMCd ('KT') ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmKt<> '' and RmPrtKey=@wCurrPrtn
			and not exists (select 'x' from Param where 
			PTyp= 'KT' and PMCd= RmKt)
		if @row> 0 
		begin
			select @errmsg= 'Karat Not Defined'
			goto ErrHndlr
		end
	
		/* Current Partition  */
		/* *** Manali 3.03 - 17/06/08 - 'L'  added */
		select @row= count('x') from inserted where RmCtg in ('G', 'P', 'S', 'L') and RmKt= '' and RmPrtKey=@wCurrPrtn
		if @row> 0 
		begin
			select @errmsg= 'Karat Not Defined'
			goto ErrHndlr
		end
	
		/****** RmZ = PMCd ('YN') ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmPrtKey=@wCurrPrtn and
			not exists (select 'x' from Param where PTyp= 'YN' and PMCd= RmZ)
		if @row> 0 
		begin
			select @errmsg= 'Specify If Customs Code (Y/N)'
			goto ErrHndlr
		end
	
		/****** If RmZ = 'N'  then 
				If RmCtg = 'A'  then  Blank 	else RmCd from RmMst (RmCtg= RmCtg, RmZ= 'Y')  
			If RmZ = 'Y'  then Blank 	******/
		
		/* Check on Current Db - added the deleted condition so that push-to-current ignores this check */
		select @row= count('x') from inserted where inserted.RmZ= 'N' and inserted.RmPrtKey=@wCurrPrtn and
			inserted.RmCtg<> 'A'  and 
			Not Exists (Select 'x' From Deleted del Where inserted.RmCd=del.RmCd and inserted.RmPrtKey<>del.RmPrtKey) and 
			Not exists (select 'x' from RmMst Rm where 
				Rm.RmCtg= inserted.RmCtg and Rm.RmZ= 'Y' and Rm.RmCd= inserted.RmZCd and Rm.RmPrtKey=inserted.RmPrtKey)
		if @row> 0 
		begin
			select @errmsg= 'Customs Code Not Defined'
			goto ErrHndlr
		end

		/* Current Partition  */
		select @row= count('x') from inserted where (RmZ= 'Y' or RmCtg= 'A') and RmPrtKey=@wCurrPrtn  
			and RmZCd<> '' 
		if @row> 0 
		begin
			select @errmsg= 'Customs Code Should Not Be Entered'
			goto ErrHndlr
		end

	/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
	/*	Date: 09/12/2003																			 */
		/****** If RmZ = 'N' And RmZCd<>'' And inserted RmZCd <> deleted RmZCd	******/
		/******		then  RmCd from RmMst (RmValidYN=)  						******/
		/* Check On Current Db - updated the deleted condition so that push-to-current ignores this check */
		Select @row= count('x') From Inserted Ins Where Ins.RmZ = 'N' And Ins.RmZCd <> '' 
				And Ins.RmPrtKey=@wCurrPrtn and 
					Ins.RmZCd <> IsNull((Select RmZCd from Deleted Del Where Del.RmCd = Ins.RmCd 
										and Del.RmPrtKey=(Case When Del.RmPrtKey<>@wCurrPrtn Then Del.RmPrtKey Else Ins.RmPrtKey End)), '') And
				Not Exists (Select 'x' from RmMst Rm Where Rm.RmCd = Ins.RmZCd And Rm.RmValidYN = 'Y' 
							And Rm.RmPrtKey=ins.RmPrtKey)
		if @row> 0
		begin
			select @errmsg= 'Customs Code Not Valid'
			goto ErrHndlr
		end
	/*  #########################################  Manoj  #########################################  */
	
		/****** RmBaseCd = valid RmCode of RmCtg 'G' or 'P' or 'S' when RmCtg= 'X' and RmZ= 'N' ******/
		/* Check On Current Db */
		/* *** Manali 3.03 - 17/06/08 - 'L'  added */
		select @row= count('x') from inserted ins where RmCtg= 'X' and RmZ= 'N' and RmPrtKey=@wCurrPrtn and
			not exists (select 'x' from RmMst Rm where Rm.RmCtg in ('G', 'P', 'S', 'L') and Rm.RmZ= 'N' 
				    and Rm.RmCd= ins.RmBaseCd and Rm.RmPrtKey=ins.RmPrtKey)
		if @row> 0 
		begin
			select @errmsg= 'Base Code For Accessories Should Be A Valid RmCode Of RmCtg ''G'' or ''P'' or ''S'' or ''L'''
			goto ErrHndlr
		end
	
		/****** RmBaseCd = Blank when RmZ= 'Y' or RmCtg= G,P,S,D,C,A,M ******/
		/* Current Partition  */
		/* *** Manali 3.03 - 17/06/08 - 'L'  added */
		select @row= count('x') from inserted ins where RmBaseCd<> '' and RmPrtKey=@wCurrPrtn and
			(RmZ= 'Y' or RmCtg in ('G', 'P', 'S', 'L', 'D', 'C', 'A', 'M')) 
		if @row> 0 
		begin
			select @errmsg= 'Base Code Should Be Blank For RmCtg= G,P,S,D,C,A,M and Custom Codes'
			goto ErrHndlr
		end
	
		/* ****** Sachin 2.13.0 - 29-05-2006 ****** */
		/****** Enter 'Q' or 'W' in Sales RmQw for all RmCategories ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmPrtKey=@wCurrPrtn and 
				not exists (select 'x' from Param where PTyp= 'LABQW' and PMCd= RmQw and PSCd= '')
		if @row> 0 
		begin
			select @errmsg= 'Mention If RM Sale Value Calculation Is By Quantity Or Weight'
			goto ErrHndlr
		end
	
		/****** Enter 'Q' or 'W' in Cost RmQw for all RmCategories ******/
		/* Current Partition  */
		select @row= count('x') from inserted where RmPrtKey=@wCurrPrtn and
				not exists (select 'x' from Param where PTyp= 'LABQW' and PMCd= RmQwCst and PSCd= '')
		if @row> 0 
		begin
			select @errmsg= 'Mention If RM Cost Value Calculation Is By Quantity Or Weight'
			goto ErrHndlr
		end

		/***** Manali 3.8.0 - AlyCd for RmCtg GPSL *****/
		select @row= count('x') from inserted Ins where RmPrtKey=@wCurrPrtn and RmCtg In ('G', 'P', 'S', 'L')
				and RmAlyCd<>'' and not exists(select 'x' from RmMst Rm where Rm.RmCtg= 'A' and Rm.RmCd=Ins.RmAlyCd and Rm.RmZ='N' and Rm.RmValidYn='Y')
		if @row> 0 
		begin
			select @errmsg= 'Invalid Alloy Code'
			goto ErrHndlr
		end

		select @row= count('x') from inserted Ins where RmPrtKey=@wCurrPrtn and RmCtg Not In ('G', 'P', 'S', 'L')
				and RmAlyCd<>'' 
		if @row> 0 
		begin
			select @errmsg= 'Alloy Code Cannot Be Entered'
			goto ErrHndlr
		end

		/* JewelSmith.2 Stone Class <>'' */
		select @row= count('x') from inserted where RmStnCls= '' and RmPrtKey=@wCurrPrtn and RmCtg in ('C', 'D')   			
		if @row> 0 
		begin
			select @errmsg= 'Stone Class Should Not Be Blank'
			goto ErrHndlr
		end
		/* JewelSmith.2 Stone Class ='' for other than DC */
		select @row= count('x') from inserted where RmStnCls<> '' and RmPrtKey=@wCurrPrtn and RmCtg not in ('C', 'D')   			
		if @row> 0 
		begin
			select @errmsg= 'Stone Class Should Be Blank'
			goto ErrHndlr
		end

		/****** RmQw = Blank if RmCtg in ('G', 'P', 'S', 'D', 'C', 'A', 'M') ******/
		/*select @row= count('x') from inserted where RmQw<> '' and RmCtg in ('G', 'P', 'S', 'D', 'C', 'A', 'M') 
	
		if @row> 0 
		begin
			select @errmsg= 'Q/W Should Be Blank For Categories G,P,S,D,C,A and M'
			goto ErrHndlr
		end*/
		
		--4 ZRmPurity cannot change if customs match exists 
		select @row= count('x') from inserted ins where ins.RmCtg In ('G', 'P', 'S', 'L','M') and ins.RmPrtKey=@wCurrPrtn and 
			Not Exists (Select 'x' from deleted del where ins.RmCd= del.RmCd and ins.RmPurityZ= del.RmPurityZ and ins.RmPrtKey=del.RmPrtKey) 
			and Exists (Select 'x' from InvRm join InvZRm on IzTc =IrTc and IzYy =IrYy and IzChr =IrChr and IzNo =IrNo 
			 where IrRmCd = RmCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Change ZPurity As RmCd Is Used In InvRm Table' 
				goto ErrHndlr
			end
		
	End

	if @mode = 'D' And @ReplAgt <> 1
	Begin

		/****** Cannot delete RmCd where RmZ= 'N' which  is in use in CustAly Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'N' and RmPrtKey=@wCurrPrtn and
			exists (select 'x' from CustAly where CaRmCd= RmCd or CaAlyCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In CustAly Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'N' which  is in use in DsgRm Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'N' and RmPrtKey=@wCurrPrtn and
			exists (select 'x' from DsgRm where DrRmCd= RmCd or DrAlyCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In DsgRm Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'Y' which  is in use in RmMst Table ******/
		/* Check On Current Db */
		select @row= count('x') from deleted del where del.RmZ= 'Y' and 
			exists (select 'x' from RmMst Rm where Rm.RmZ= 'N' and Rm.RmZCd= del.RmCd and Rm.RmPrtKey=del.RmPrtKey)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Custom Code In Use In RmMst Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'N' which  is in use in RmRt Table ******/
		/* Check On Current Db */
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from RmRt where RrCd= RmCd and RrPrtKey=RmPrtKey)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In RmRt Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'N' which  is in use in OrdRm Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from OrdRm where OrRmCd= RmCd or OrAlyCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In OrdRm Table'
			goto ErrHndlr
		end

		/****** Manali 3.10.0 - Cannot delete RmCd where RmZ= 'N' which  is in use in MultiPrcQtRm Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from MultiPrcQtRm where MqrRmCd= RmCd or MqrAlyCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In MultiPrcQtRm Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'N' which  is in use in OrdTgt Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from OrdTgt where OtRmCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In OrdTgt Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'N' which is in use in FgRm Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from FgRm where FrRmCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In FgRm Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'N' which is in use in SYyMm Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from SYyMm where SRmCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In SYyMm Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'N' which is in use in Txnd Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from Txnd where TdRmCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In Txnd Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'N' which is in use in Grade Table ******/
		/* Partition Check Not required */
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from Grade where GrRmCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code Is Used In Grade Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'Y' which is in use in TxndZ Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'Y' and 
			exists (select 'x' from TxndZ where TzRmZCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Customs Code In Use In TxndZ Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where RmZ= 'Y' which is in use in InvRm Table ******/
		/* Check On Complete Db */
		select @row= count('x') from deleted del where RmZ= 'Y' and 
			exists (select 'x' from InvRm where IrRmZCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Customs Code In Use In InvRm Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where which is in use in RmAlyCmp Table ******/
		/* Partition Check Not Required */
		select @row= count('x') from deleted del where 
			exists (select 'x' from RmAlyCmp where RaCd= RmCd or RaCmpCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm/ Customs Code In Use In RmAlyCmp Table'
			goto ErrHndlr
		end
		
		/*VK.5*/		
		/****** Cannot delete RmCd which is in use in CertDet Table ******/		
		select @row= count('x') from deleted del where RmZ= 'N' and 
			exists (select 'x' from CertDet where CdRmCd= RmCd)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use In CertDet Table'
			goto ErrHndlr
		end

		/****** Cannot delete RmCd where which is in use as Base Code in RmMst Table ******/
		/* Check On Current Db */
		select @row= count('x') from deleted del where del.RmZ= 'N' and del.RmCtg in ('G', 'P', 'S', 'L') and 
			exists (select 'x' from RmMst Rm where Rm.RmCtg= 'X' and Rm.RmZ= 'N' 
						and Rm.RmBaseCd= del.RmCd and Rm.RmPrtKey=del.RmPrtKey)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use As Base Code In RmMst Table'
			goto ErrHndlr
		end

		select @row= count('x') from deleted del where del.RmZ= 'N' and del.RmCtg in ('A') and 
			exists (select 'x' from RmMst Rm where Rm.RmCtg In ('G', 'P', 'S', 'L') and Rm.RmZ= 'N' 
						and Rm.RmAlyCd= del.RmCd and Rm.RmPrtKey=del.RmPrtKey)
		if @row> 0 
		begin
			select @errmsg= 'Cannot Delete As Rm Code In Use As Alloy Code In RmMst Table'
			goto ErrHndlr
		end

	End

	if @SYS= 'XEMR' 
	/* Partition Checks Not done On 'XEMR' */
	BEGIN

		if @mode = 'D' 
		Begin
			/****** Cannot delete RmCd where RmZ= 'N' which  is in use in xTxnRm Table ******/
			select @row= count('x') from deleted del where RmZ= 'N' and 
				exists (select 'x' from xTxnRm where xTrRmCd= RmCd
					or xTrRmCtg= 'A' and  xTrAlyCd= RmCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Rm Code In Use In xTxnRm Table'
				goto ErrHndlr
			end	
	
			/****** Cannot delete RmCd where RmZ= 'N' which  is in use in xTxnTgt Table ******/
			select @row= count('x') from deleted del where RmZ= 'N' and 
				exists (select 'x' from xTxnTgt where xTtRmCd= RmCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Rm Code In Use In xTxnTgt Table'
				goto ErrHndlr
			end

			/****** Cannot delete RmCd where RmZ= 'N' which is in use in xStkYyMm Table ******/	
			select @row= count('x') from deleted del where RmZ= 'N' and 
				exists (select 'x' from xStkYyMm where xSyRmCd= RmCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Rm Code In Use In xStkYyMm Table'
				goto ErrHndlr
			end
		End
	END

return
ErrHndlr:
	begin
		select @errmsg='|'+@errmsg
		raiserror(@errmsg,16,1)
		rollback transaction 
	end
END

