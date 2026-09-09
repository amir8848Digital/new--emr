
ALTER TRIGGER [dbo].[Param_INUPDL] ON [dbo].[Param] 
       FOR INSERT,UPDATE,DELETE NOT FOR REPLICATION
AS
BEGIN
	/****** Trigger Version - 3.11.1 ******/
	DECLARE @ROW INT
	DECLARE @mode char(1)
	DECLARE @errmsg varchar(255)
	DECLARE @PTyp varchar(8)
	DECLARE @wHCd VARCHAR(8)
	DECLARE @SYS VARCHAR(5)
 	DECLARE @ReplAgt BIT  

	IF @@ROWCOUNT= 0
		RETURN
	If Exists (Select 'x' From TrgDisa Where TdTblNm = 'Param')
		RETURN

	Set @wHCd = 'ZSELF'	
	Set @SYS= IsNull((Select PMCd from Param where PTyp= 'SYS' and PMCd= 'XEMR' and PSCd= ''), '') 

	/****** Determining whether the trigger has been called in an insert, update or delete mode ******/
	/****** If there no records in inserted the the mode is 'Delete' 			    ******/
	/****** If there no records in deleted the the mode is 'Insert' 		            ******/
	/****** If there records in both inserted and deleted the the mode is 'Update' 		    ******/
	Select @mode=case (select count('x') from inserted) when 0 then 'D'
		Else Case (select count('x') from deleted) when 0 then 'A' else 'C' end
       	End 

	/* **** Checking Whether a Normal Transaction or Replicated One **** */
	Set @ReplAgt=Convert(BIT,SessionProperty('Replication_Agent'))

	/*** Storing PTyp In The Variable @PTyp ***/
	select @PTyp= case  when @mode in ('C', 'A') then (Select max(PTyp) from inserted)  else (Select max(PTyp) from deleted) end  

	If  (@ReplAgt<>1)
	Begin
	 
		/******Cannot Insert/Update/Delete if Crm database /Emr database with sdMst='N' ****/
		if @mode in ('A','C','D')		
		Begin 
		If   Exists ( Select 'x' From SysDb where SysId='EMR' and SdMst='N' )  Or  
		     Exists ( Select 'x' From SysDb where SysId='CRM' and SysDb='WEB' ) 
			Begin 
				select @errmsg= 'Cannot Insert /Update /Delete ' 
				goto ErrHndlr 
			end 
		End 
		
		if @mode in ('A','C')
		Begin
			If @PTyp= 'SUBPRC' 
			Begin 
				/****** PMCd = valid Process Loc ******/
				/*uma **** P replaced by R ***/
				select @row= count('x') from inserted where not exists (select 'x' from Loc where LocTyp= 'R' and LocCd= PMCd)
				if @row>0 
				begin
					select @errmsg= 'Process Not Defined'
					goto ErrHndlr
				end
			End
			
			/*********Bhavna PMCD = Valid proc loc for rework ******/
			/**** uma P to R******/
			If @PTyp= 'RWRK' 
			Begin 
				/****** PMCd = valid Process Loc ******/
				select @row= count('x') from inserted where not exists (select 'x' from Loc where LocTyp= 'R' and LocCd= PMCd)
				if @row>0 
				begin
					select @errmsg= 'Process Not Defined'
					goto ErrHndlr
				end
			End
			/*********Bhavna PMCD = Valid proc loc for rework ******/
		End
	End  

	Set DateFormat DMY 

	/***Geeta****From CRM 2.12 Onwards 
	NOTE::	***PTYP  Should be Unique in Param,vParam & cParam*****************/ 

 	/******** Records Of Only One PTyp Can Be Inserted At A Time ********/ 
	select @Row= Count(distinct(PTyp)) from inserted 
	If @Row > 1
	Begin 
		Select @ErrMsg='Parameter Of Only One Type Can Be Acted On'
		GoTo ErrHndlr
	End

	/******** Records Of Only One PTyp Can Be Deleted At A Time ********/ 
	select @Row= Count(distinct(PTyp)) from deleted 
	If @Row > 1 
	Begin 
		Select @ErrMsg='Parameter Of Only One Type Can Be Acted On'
		GoTo ErrHndlr
	End

	/******** If The Parameter Is A System Parameter Then No Add, Change Or Delete Allowed For The Param Entry *********/
	/* (Bef 3.01)
	select @Row=Count('x') from Param Pm where (@PTyp= 'PTYP') or 
		(Pm.PTyp= 'PTYP' and Pm.PMCd= @PTyp and Pm.PValue= 'M')
	(Bef 3.01) */

	/*** (Jen 3.01) ***/

	select @Row= Count('x') from Param Pm where (@PTyp= 'PTYP') or 
		(Pm.PTyp= 'PTYP' and Pm.PMCd= @PTyp and Pm.PValue= 'M' and @PTyp<> 'VER')
	/*** (Jen 3.01) ***/
	If @Row > 0
	begin 
		Select @ErrMsg= 'Cannot Add/Change/Delete Microway Parameter - '+ @PTyp
		GoTo ErrHndlr
	end		

	/*** (Jen 3.01) ***/
	select @Row= Count('x') from Param Pm where @mode in ('A','D') and 
		Pm.PTyp= 'PTYP' and Pm.PMCd= @PTyp and Pm.PValue= 'M' and @PTyp= 'VER' 
	If @Row > 0
	begin 
		Select @ErrMsg= 'Cannot Add/Delete Microway Parameter - '+ @PTyp
		GoTo ErrHndlr
	end		
	/*** (Jen 3.01) ***/
	
	select @Row= Count('x') from Inserted where @mode ='A' and PTyp= 'LOCKDT' 
	If @Row > 0
	begin 
		Select @ErrMsg= 'Cannot Add  LOCKDT Parameter '
		GoTo ErrHndlr
	end	

	select @Row= Count('x') from Deleted where @mode ='D' and PTyp= 'LOCKDT' 
	If @Row > 0
	begin 
		Select @ErrMsg= 'Cannot Delete  LOCKDT Parameter '
		GoTo ErrHndlr
	end		

	select @Row= Count('x') from Inserted where @mode ='A' and PTyp= 'DMMATX' 
	If @Row > 0
	begin 
		Select @ErrMsg= 'Cannot Add Records to Design Matrix Parameter '
		GoTo ErrHndlr
	end	
			
		
 	/******** Key Cannot Be Changed  (i.e. PTyp, PMCd and PSCd cannot be changed) ********/ 
 	if @mode = 'C'
 	Begin
  		select @Row= Count('x') from deleted del where
    		not exists (Select 'x' from inserted ins where ins.PTyp= del.PTyp and ins.PMCd= del.PmCd and ins.PSCd= del.PSCd)
  		If @Row > 0
	  	begin 
			Select @ErrMsg='Cannot Change Param Key'
			GoTo ErrHndlr
		end
	End

	if @mode in ('A','C')
	BEGIN
		/*** Only a valid PTyp can be entered (i.e. a valid PMCd from Param where PTyp= 'PTYP') ***/
		select @row=count('x') from inserted ins where ins.PTyp<> 'PTYP' and 
		not exists (select 'x' from Param Prm where Prm.PTyp='PTYP' and Prm.PMCd= ins.PTyp)
		if @row>0 
		begin
			select @errmsg='Parameter Type Not Defined'
			goto ErrHndlr
		end
		/*** PMCd cannot be Blank ***/
		select @row= count('x') from inserted where PMCd= ''
		if @row>0 
		begin
	   		select @errmsg='Main Code Cannot Be Blank'
   			goto ErrHndlr
  		end

		/***  PDesc cannot be blank ***/
		/* *** Jay 3.2.0 [Stmp Instr] */
		/**** Manali 3.8.1 - SETCD started ****/
		/* IG.44 - BMRsnCd PDesc = Blank ****/
		--6.1B-26 USRCMCD ADDED
		select @row= count('x') from inserted where PDesc= '' and @PTyp not in ('INVDESC', 'STMPINST', 'SETCD','DMMATX','BMRSNCD','PYMTRM','USRCMCD')		
		if @row>0 
		begin
			select @errmsg= 'Description Cannot Be Blank'
			goto ErrHndlr
  		end

		if @PTyp= 'CPYRT' 
		Begin 
			/****** PSCd= Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row> 0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue= Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum= 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end

/*** ###
			select @row=count('x') from inserted where PDesc225<> '' and 
			if @row>0 
			begin
				select @errmsg= 'Raw Material Category In Sequence Not Defined'
				goto ErrHndlr
			end
***/
			select @row= count('x') from inserted where PDesc225= '' 
			if @row> 0 
			begin
				select @errmsg= 'Raw Material Category Sequence Should Be Entered'
				goto ErrHndlr
			end

			/****** PValue1= Blank ******/		
			select @row= count('x') from inserted where PValue1<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1= 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end

/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */

			/* Bef 3.03 - Following check no more reqd for PmCd= 'GPSDCXAM' */
			/****** Cannot add or change the Record of PmCd= 'GPSDCXAM' ******/
			/*
			select @row= count('x') from inserted where PMCd= 'GPSDCXAM'
			if @row> 0 
			begin
				select @errmsg= 'Cannot Add Or Change The Record ''GPSDCXAM'' '
				goto ErrHndlr
			end
			*/

			/* Manali 3.03 - 12/07/08 */
			/****** Cannot add or change the Record of PmCd= 'GPSLDCXM' ******/
			select @row= count('x') from inserted where PMCd= 'GPSLDCXM'
			if @row> 0 
			begin
				select @errmsg= 'Cannot Add Or Change The Record ''GPSLDCXM'' '
				goto ErrHndlr
			end
			/* Manali 3.03 - 12/07/08 */
		End

		if @PTyp= 'CURNCY' 
		Begin 
			/****** PMCd= Not Blank ******/
			select @row= count('x') from inserted where PMCd= ''
			if @row> 0 
			begin
				select @errmsg= 'Main Code Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row> 0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/* Zubin 213 (Bef 213) */ 
			/****** PValue= Blank ******/
/*			select @row= count('x') from inserted where PValue<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
*/ 
			/* Zubin 213 */ 
			select @row= count('x') from inserted where PValue= '' 
			if @row> 0 
			begin
				select @errmsg= 'Currency Unit Description Cannot Be Blank'
				goto ErrHndlr
			end

			/****** Round(PNum, 3)> 0 ******/
			select @row= count('x') from inserted where Round(PNum, 3)<= 0
			if @row> 0 
			begin
				select @errmsg= 'Currency Conversion Rate Should Be > Zero'
				goto ErrHndlr
			end
/* COMMENTED TEMPORARILY (WILL BE INCORPORATED LATER) */ 
			/* Zubin 213 */ 
/*			select @row= count('x') from inserted where PNum<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
*/
			/****** PDesc225= Blank ******/
			select @row= count('x') from inserted where PDesc225<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/* Zubin 213 (Bef 213) */ 
			/****** PValue1= Blank ******/		
/*			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
*/
			/* Zubin 213 */ 
			select @row=count('x') from inserted where PValue1= ''
			if @row>0 
			begin
				select @errmsg= 'Decimal Unit Description'
				goto ErrHndlr
			end
			/****** PNum1= 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			/**** Manali 3.9.0 - SB File Cur Cd Mapped In this field ***/
			/*
			select @row= count('x') from inserted where PValue2<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			*/
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */

			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End 

		if @PTyp= 'DAANACD' 
		Begin 
			/****** PMCd= Valid Dsg Ana Fld where PValue= 'Y' ******/ 
			select @row= count('x') from inserted ins where 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'DAANAFLD' and Prm.PMCd= ins.PMCd and Prm.PValue= 'Y') 
			if @row> 0 
			begin
				select @errmsg= 'Design Analysis Field Not Defined'
				goto ErrHndlr
			end
			/****** PSCd<> Blank ******/ 
			select @row= count('x') from inserted where PSCd= ''
			if @row> 0 
			begin
				select @errmsg= 'Design Analysis Code Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue= Blank ******/
			select @row= count('x') from inserted where PValue<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where Str(PNum, 16, 0)<> Str(0, 16, 0)
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1= Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */

		End

		if @PTyp= 'RMSCTG'		/*	--|| 4.1.4 ||-- */
		Begin 
			if @mode in ('C') 
			begin
				Select @row= count('x') from inserted ins where 
					Exists (Select 'x' from deleted del where del.PMCd= ins.PMCd and del.PSCd= ins.PSCd and del.PValue<>ins.PValue and del.PValue='BS') 
					And Exists (Select 'x' from RmRt RrT
							    Where RrT.RrCtg=ins.PMCd And RrT.RrSCtg=ins.PSCd And RrT.RrTcTyp In ('ILC','NAS','RM','RSZ','STW') 									
									And not exists (Select * from RmRt Rr where Rr.RrTcTyp= 'BS' and
				   									Rr.RrCmCtg=RrT.RrCmCtg and Rr.RrCmCd= RrT.RrCmCd and 
													Rr.RrCmCurCd='' and Rr.RrCtg= RrT.RrCtg and
   													Rr.RrSCtg= RrT.RrSCtg and Rr.RrSTWGrp= '' and 
													Rr.RrCd= '' and Rr.RrDmCtg= '' and Rr.RrDmCd=RrT.RrDmCd and 
													Rr.RrLsCtg= '' and Rr.RrLabMCd= '')
									And (Select count('x') from RmRt Rt 
											Where Rt.RrTcTyp= RrT.RrTcTyp and Rt.RrCmCtg= RrT.RrCmCtg and Rt.RrCmCd= RrT.RrCmCd and 
												  Rt.RrCmCurCd= RrT.RrCmCurCd and Rt.RrCtg= RrT.RrCtg and Rt.RrSCtg= RrT.RrSCtg and 
												  Rt.RrSTWGrp= RrT.RrSTWGrp and Rt.RrCd= RrT.RrCd and Rt.RrDmCtg= RrT.RrDmCtg and 
												  Rt.RrDmCd=RrT.RrDmCd and Rt.RrLsCtg= RrT.RrLsCtg and Rt.RrLabMCd= RrT.RrLabMCd
													and ((Rt.RrFrLn<= RrT.RrFrLn and Rt.RrToLn>= RrT.RrFrLn)
													or 	(Rt.RrFrLn<= RrT.RrToLn and Rt.RrToLn>= RrT.RrToLn)
													or  (RrT.RrFrLn<= Rt.RrFrLn and RrT.RrToLn>= Rt.RrToLn))) > 1) 
				if @row> 0 
				begin 
					select @errmsg= 'Cannot Chg ''Rate By'' as Duplicate Range Size Exists in one of the RmRt Types ''ILC'',''NAS'',''RM'',''RSZ'' or ''STW'' ' 
					goto ErrHndlr
				end 
			end			
		End

		if @PTyp= 'DAANAFLD' 
		Begin 
			if @mode in ('C') 
			begin
				/****** Cannot Change PValue when DsgAna Rec Exists *****/
				select @row= count('x') from inserted ins where 
					not exists(Select 'x' from deleted del where del.PMCd= ins.PMCd 
						and del.PSCd= ins.PSCd and del.PValue= ins.PValue) 
					and exists(Select DaAnaSr from DsgAna where DaAnaSr= ins.PMCd) 
				if @row> 0 
				begin 
					select @errmsg= 'Cannot Change Defn Of Ana Field When Rec Exists In DsgAna' 
					goto ErrHndlr
				end 
			end			

			/****** PSCd= Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue= valid PMCd('YN') ******/
			select @row=count('x') from inserted ins where 
			not exists (select 'x' from Param Prm where Prm.PTyp= 'YN' and Prm.PMCd= ins.PValue)
			if @row> 0 
			begin
				select @errmsg= 'Specify Whether This Is An Ana Cd Fld Or Not (''Y'' For Cd/ ''N'' For Desc)'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where Str(PNum, 16, 0)<> Str(0, 16, 0)
			if @row> 0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 <> Blank ******/
			select @row= count('x') from inserted where PDesc225= ''
			if @row> 0 
			begin
				select @errmsg= 'Short Description Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue1= valid PMCd('YN') ******/
			select @row= count('x') from inserted ins where 
			not exists (select 'x' from Param Prm where Prm.PTyp= 'YN' and Prm.PMCd= ins.PValue1) 
			if @row> 0 
			begin
				select @errmsg= 'Specify Whether This Field Is Compulsory Or Not'
				goto ErrHndlr
			end
			/****** If PValue= 'Y' then PValue1= 'Y' ******/
			select @row= count('x') from inserted ins where PValue= 'Y' and PValue1<> 'Y' 
			if @row> 0 
			begin
				select @errmsg= 'Design Analysis Codes Have To Be Compulsory'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row> 0 
			begin
				select @errmsg='PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		/* Jenny Color (02/04/05) */
		if @PTyp= 'DEFCLR' 
		Begin 
			/****** Only One Record Of 'DEFCLR' Is Allowed
				 is allowed ******/
			select @row= count('x') from Param where PTyp= 'DEFCLR' 
			if @row> 1
			begin
				select @errmsg= 'Only One Record Allowed For This Param Type'
				goto ErrHndlr
			end
			/****** PMCd= 'DEFCLR' ******/
			select @row= count('x') from inserted where PMCd<> 'DEFCLR'
			if @row> 0 
			begin
				select @errmsg= 'Main Code Should Be DEFCLR'
				goto ErrHndlr
			end
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row> 0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = valid CsCd from ClrSet ******/
			/*** (Bef 2.13)
			select @row= count('x') from inserted where 
				Not Exists (Select 'x' From ClrSet Where CsCd= PValue) 
			***/
			/*** (Jen 2.13) ***/
			select @row= count('x') from inserted where PValue<> '' and 
				Not Exists (Select 'x' From ClrSet Where CsCd= PValue) 
			/*** (Jen 2.13) ***/
			if @row> 0 
			begin
				select @errmsg= 'Color Code Not Defined'
				goto ErrHndlr
			end 
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row> 0  
			begin
				select @errmsg='PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row> 0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		/* Jenny Color (02/04/05) */


		if @PTyp= 'DMCTG' 
		Begin 
			/****** PSCd= Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue= '' or valid Design Unit Of Measurement ******/
			select @row=count('x') from inserted ins where ins.PValue<> '' and 
			not exists (select 'x' from Param Prm where Prm.PTyp= 'DMUOM' and Prm.PMCd= ins.PValue)
			if @row>0 
			begin
				select @errmsg= 'Design Unit Of Measurement Not Defined'
				goto ErrHndlr
			end
			/****** PNum > 0 ******/
			select @row=count('x') from inserted where PNum= 0
			if @row>0 
			begin
				select @errmsg= 'Bag Opening Pieces Should Be Greater Than Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			/*** PrcsSeq adding in PDesc225*********/
			/*select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end*/
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg='PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
	/*	Manoj 2.10.0	*/
			/****** PValidYN = Blank ******/
/*			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
*/
			select @row= count('x') from inserted ins where
				Not Exists (Select 'x' From Param Where PTYp = 'YN' And PMCD = ins.PValidYN)
			if @row>0
			begin
				select @errmsg= 'PValidYN Not Defined'
				goto ErrHndlr
			end
	/*	Manoj 2.10.0	*/
/*  #########################################  Manoj  #########################################  */
		End


		/* CMGRP - Sachin 4.1.0.0 */
		if @PTyp in ('CMGRP','SKL') 
		Begin 
			/****** PSCd= Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end

			/****** PValue= '' ******/
			select @row=count('x') from inserted ins where ins.PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end

			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end

			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg='PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			select @row= count('x') from inserted ins where PValidYN<>''				
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End

		if @PTyp= 'CMTYP' 
		Begin 
			/****** PSCd= Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end

			/****** PValue= '' ******/
			select @row=count('x') from inserted ins where ins.PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end

			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end

			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg='PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			select @row= count('x') from inserted ins where PValidYN<>''				
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		/* Sachin 4.1.0.0 - Changes End */

		/* DMMATX - Sachin 4.1.2.0 */
		if @PTyp= 'DMMATX' 
		Begin 
			/****** PSCd= Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end

			/****** PValue= '' ******/
			select @row=count('x') from inserted ins where ins.PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end

			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end

			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg='PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			select @row= count('x') from inserted ins where PValidYN<>''				
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End

		if @PTyp= 'DSGCPY'
		Begin 
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted ins where PDesc225<> '' and 
				not exists(Select 'x' from Param P where P.PTyp= 'TCTYP' 
					and P.PMCd = ins.PDesc225 and P.PValue1= 'OM') 
			if @row> 0 
			begin
				select @errmsg= 'TcType Not Defined'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End

		if @PTyp= 'GLDAS' 
		Begin 
			/****** PSCd= Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue= Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum= 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225<> Blank ******/
			select @row= count('x') from inserted where PDesc225= '' 
			if @row> 0 
			begin
				select @errmsg= 'Raw Material Category Sequence Should Be Entered'
				goto ErrHndlr
			end

			/*** (Jen 2.14) ***/
			/*** Cannot Change PDesc225 when record exists in OrdMst, OrdDsg, InvDsg ***/
			/*** Not checking for Head, CustMst ***/
	  		select @Row= Count('x') from deleted del where 
	    		not exists (Select 'x' from inserted ins where ins.PTyp= del.PTyp and ins.PMCd= del.PmCd 
					and ins.PSCd= del.PSCd and ins.PDesc225= del.PDesc225) 
			And (	exists (Select 'x' from OrdMst where OmGldAs= del.PmCd or OmLabAs= del.PmCd) 
				or exists (Select 'x' from OrdDsg where OdGldAs= del.PmCd or OdLabAs= del.PmCd) 
				or exists (Select 'x' from InvDsg where IdGldAs= del.PmCd or IdLabAs= del.PmCd) 
			    )
			If @Row> 0
			begin 
				Select @ErrMsg= 'Cannot Change Rm Category Sequence As Record In Use In Order/ Invoice'
				GoTo ErrHndlr
			end		
			/*** (Jen 2.14) ***/


/*** ###
			select @row=count('x') from inserted where PDesc225<> '' and 
			if @row>0 
			begin
				select @errmsg= 'Raw Material Category In Sequence Not Defined'
				goto ErrHndlr
			end
***/
			/****** PValue1= Blank ******/		
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1= 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg='PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg='PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		/* *** Jay 3.2.0 [Stmp Instr] */
		/* IG.44 -BMRsnCd addition is added ****/
		if @PTyp in ('INVDESC','STMPINST','BMRSNCD','PYMTRM')
		/* *** Jay 3.2.0 [Stmp Instr] */
		Begin 
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end

			/****** PDesc225 = Not Blank ******/
			 /* *** Jay 3.2.0 [Stmp Instr] */
			select @row=count('x') from inserted where PDesc225= '' 
			if @row>0 
			begin
				if @PTyp = 'INVDESC' select @errmsg= 'Enter Detail Product Description'
				else if @PTyp = 'STMPINST' select @errmsg= 'Enter Stemping Instruction Can Not Be Blank'
				else if @PTyp = 'PYMTRM' select @errmsg= 'Payment Term Description Can Not Be Blank'
				else if @PTyp = 'BMRSNCD' select @errmsg= 'B\M Reason Description Can Not Be Blank'
				goto ErrHndlr
			end
			 /* *** Jay 3.2.0 [Stmp Instr] */

			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
			select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'IWTFRORD' 
		Begin 
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue= Blank ******/
			select @row= count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum= 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225<> Blank ******/	
			select @row= count('x') from inserted where PDesc225= '' 
			if @row> 0 
			begin
				select @errmsg= 'Raw Material Category Sequence Should Be Entered'
				goto ErrHndlr
			end
/*** ###
			select @row= count('x') from inserted where PDesc225<> '' and 
			if @row>0 
			begin
				select @errmsg= 'Raw Material Category In Sequence Not Defined'
				goto ErrHndlr
			end
***/
			/****** PValue1= Blank ******/		
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1= 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'LABGRP' 
		Begin 
			/****** PMCd= A report menu option like 'DSGCRD', 'ORDAVG', 'ORDDET' ******/			
			/****** PSCd= Not Blank ******/
			select @row=count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Labour Group Option'
				goto ErrHndlr
			end
			/****** PValue= Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum= 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225= Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1= Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

				if @PTyp= 'LABLOGIC' 
		Begin 
			/****** PMCd = Valid Labour Main Code ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'LABMCD' 
				and Prm.PMCd= ins.PMCd)
			if @row>0 
			begin
				select @errmsg= 'Labour Main Code Not Defined'
				goto ErrHndlr
			end
			/****** PSCd= Not Blank ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'Labour Logic Code Cannot Be Blank'
				goto ErrHndlr
			end

				
			/****** Manali 3.5.0 - If PMCd= ('DHC','CHC') then PSCd= '-' ******/
			/* ***** Manali - This Check not required ****** */
/*			select @row= count('x') from inserted where PMCd In  ('DHC', 'CHC') and PSCd<> '-'
			if @row> 0 
			begin
				select @errmsg= 'Labour Sub Code Should Be ''-'' For ''DHC'' and ''CHC'' '
				goto ErrHndlr
			end
*/
			/****** PValue = 'Q' or 'W' ******/

			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Value Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225= Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end			/****** PValue1 <> Blank ******/
			select @row= count('x') from inserted where PValue1= ''
			if @row>0 
			begin
				select @errmsg= 'Print Description Should Not Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end

/*  #########################################  Manoj  #########################################  */
		End


		if @PTyp= 'LABSCD' 
		Begin 
			/****** PMCd = Valid Labour Main Code ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'LABMCD' 
				and Prm.PMCd= ins.PMCd)
			if @row>0 
			begin
				select @errmsg= 'Labour Main Code Not Defined'
				goto ErrHndlr
			end
			/****** PSCd= Not Blank ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'Labour Sub Code Cannot Be Blank'
				goto ErrHndlr
			end

			/****** If PMCd= 'ZBRK' then PSCd= '-' ******/
			select @row= count('x') from inserted where PMCd= 'ZBRK' and PSCd<> '-'
			if @row> 0 
			begin
				select @errmsg= 'Labour Sub Code Should Be ''-'' For ''ZBRK'' '
				goto ErrHndlr
			end
				
			/****** Manali 3.5.0 - If PMCd= ('DHC','CHC') then PSCd= '-' ******/
			/* ***** Manali - This Check not required ****** */
/*			select @row= count('x') from inserted where PMCd In  ('DHC', 'CHC') and PSCd<> '-'
			if @row> 0 
			begin
				select @errmsg= 'Labour Sub Code Should Be ''-'' For ''DHC'' and ''CHC'' '
				goto ErrHndlr
			end
*/
			/****** PValue = 'Q' or 'W' ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'LABQW' 
			       and Prm.PMCd= ins.PValue)
			if @row>0 
			begin
				select @errmsg= 'Labour Q / W Not Defined'
				goto ErrHndlr
			end

			/****** Manali 3.5.0 - If PMCd= ('DHC','CHC') then PValue = 'Q' ******/
			select @row= count('x') from inserted where PMCd In  ('DHC', 'CHC') and PValue<> 'Q'
			if @row> 0 
			begin
				select @errmsg= 'Labour Q/W Should Be ''Q'' For ''DHC'' and ''CHC'' '
				goto ErrHndlr
			end

			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Value Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225= Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end			/****** PValue1 <> Blank ******/
			select @row= count('x') from inserted where PValue1= ''
			if @row>0 
			begin
				select @errmsg= 'Print Description Should Not Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> '' and pMcd not in ('GHC','CDW')
 			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN <> Blank & PMCD(YN)******/
			select @row= count('x') from inserted ins where
				Not Exists (Select 'x' From Param Where PTYp = 'YN' And PMCD = ins.PValidYN)
			if @row>0
			begin
				select @errmsg= 'PValidYN Not Defined'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		/* ****** Manali 3.8.0 - MELTR ***** */
		if @PTyp = 'METLR'
		Begin 
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
		
			Select @row= count('x') from inserted ins 
				Join Param P1 On P1.PTyp='USR' and ','+ins.PDesc225+',' like '%,'+P1.PMCd+',%'  
				and (Select Count('x') From Param P2 where P2.PTyp=ins.PTyp and ','+P2.PDesc225+',' like '%,'+P1.PMCd+',%')>1  
			if @row>0 
			begin
				select @errmsg= 'Same User Cannot Be Assigned Against Multiple Metler Codes'
				goto ErrHndlr
			end

			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		
		End


	/******************* urmi **********************************************************/
		if @PTyp = 'LOCKDT'
		Begin 
			/****** Only One Record Of Each PMCD ('LOCKDT' & 'FUTUREDT')
				 is allowed ******/
	/*		select @row= count('x') from Param where PTyp= 'LOCKDT' 
					and PMCd Not In ('LOCKDT', 'FUTUREDT')
			if @row> 0
			begin
				select @errmsg= 'PMCd Should Be Either ''LOCKDT'' Or ''FUTUREDT'' '
				goto ErrHndlr
			end
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
*/
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
		/*	select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
		*/
			/****** PDesc225 = a valid date ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			
			/****** PDesc225 = a valid date ******/
			/*
			select @row= count('x') from inserted 
				where Cast(PDesc225 as smalldatetime)< 
						(Select Cast(PDesc225 as smalldatetime) From Param 
								Where Ptyp='PURGEDT' And PMCd='PURGEDT')
			if @row>0 
			begin
				select @errmsg= 'Lock Date Cannot Be Less Than Purge Dt'
				goto ErrHndlr
			end
			*/

			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
/*			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
*/
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

	/******************* urmi **********************************************************/
	
		if @PTyp in ('ORDDSGGR', 'ORDLABGR', 'ORDRMGR', 'TBAGBSGR', 'TBAGVSGR', 'INBAGGR', 'INGRPRCP', 'INDET', 'INGR', 'FGBAGGR')
		Begin 
			/******** PSCd = Blank ********/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/******** PValue = Blank ********/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/******** PDesc225 = Not Blank ********/
			/******** The sequence of fields on which the grouping has to be done is mentioned here (The value in this field is actually 
				not used because thay are anyway hard coded in the places where they are used ********/
			select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Group By Fields'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
				begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

/* Zubin 213 */
/*		if @PTyp= 'OUNCE' */ 
		if @PTyp= 'GMFACTOR'  
		Begin 
			/* Zubin 213 PSCd = Valid currency */ 
			select @row= count('x') from inserted ins where 
				Not Exists (Select 'x' from Param Prm where Prm.PTyp= 'CURNCY' and Prm.PMCd= ins.PSCd) 
			if @row> 0 
			begin 
				select @errmsg= 'Currency Code Not Defined' 
				goto ErrHndlr 
			end 
			/****** PSCd = Blank (Zubin 213 [below chk not required]) ******/
/*			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
*/
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum <> 0  ******/
			select @row= count('x') from inserted where PNum= 0
			if @row>0 
			begin
				select @errmsg= 'Value Should Be Entered'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> '' 
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp =  'PARENTCD'
		Begin 
			/****** Cannot Change RmCtg, RmSCtg of Parent Code *****/
			if @mode in ('C') 
			begin
				select @row= count('x') from inserted ins where 
					not exists(Select 'x' from deleted del where del.PMCd= ins.PMCd 
						and del.PValue= ins.PValue and del.PValue1= ins.PValue1) 
				if @row> 0 
				begin 
					select @errmsg= 'Cannot Change Rm Ctg/ SubCtg Of Param Type ''PARENTCD'' ' 
					goto ErrHndlr
				end 
			end			
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row> 0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = valid PMCd('RMCTG') ******/
			select @row= count('x') from inserted ins where PValue not in ('D', 'C') 
				/* not exists(Select 'x' from Param Pm where Pm.PTyp= 'RMCTG' and Pm.PMCd= ins.PValue) */
			if @row> 0 
			begin
				select @errmsg= 'Rm Category Not Defined'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = valid PSCd('RMSCTG') ******/
			select @row= count('x') from inserted ins where 
				not exists(Select 'x' from Param Pm where Pm.PTyp= 'RMSCTG' and Pm.PMCd= ins.PValue and Pm.PSCd= ins.PValue1) 
			if @row> 0 
			begin
				select @errmsg= 'Rm Sub Category Not Defined'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		/*** (Jen 3.01.1) ***/
		if @PTyp= 'REMAKE' 
		Begin 
			/****** PSCd= Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PDesc225= Blank ******/
			select @row= count('x') from inserted where PDesc225<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end

			/****** PNum= 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end

			/****** PValue<> Blank ******/
			select @row= count('x') from inserted where PValue= '' 
			if @row> 0 
			begin
				select @errmsg= 'Raw Material Category Sequence Should Be Entered'
				goto ErrHndlr
			end

			/*** Cannot Change PValue when record exists in InvHd ***/
	  		select @Row= Count('x') from deleted del where 
	    		not exists (Select 'x' from inserted ins where ins.PTyp= del.PTyp and ins.PMCd= del.PmCd 
							and ins.PSCd= del.PSCd and ins.PValue= del.PValue) 
					and exists (Select 'x' from InvHd where InRemakingCtgs= del.PmCd) 
			If @Row> 0
			begin 
				Select @ErrMsg= 'Cannot Change Rm Category Sequence As Record In Use In Invoice'
				GoTo ErrHndlr
			end		

			/****** PValue1= Blank ******/		
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1= 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg='PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2= Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg='PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		/*** (Jen 3.01.1) ***/


/*************************Geeta********************2.10************************/
		if @PTyp =  'STNCLS'
		Begin 
			/****** Cannot Change RmCtg  of Stone Class *****/
			if @mode in ('C') 
			begin
				select @row= count('x') from inserted ins where 
					not exists(Select 'x' from deleted del where del.PMCd= ins.PMCd 
						and del.PSCd= ins.PSCd) 
				if @row> 0 
				begin 
					select @errmsg= 'Cannot Change Rm Ctg  Of Param Type ''STNCLS'' ' 
					goto ErrHndlr
				end 
			end			
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row> 0 
			begin
				select @errmsg= 'Code Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue = valid PMCd('RMCTG') ******/
			/*** (Bef 2.12)
			select @row= count('x') from inserted ins where PMCD not in ('D', 'C') 
			if @row> 0 
			begin
				select @errmsg= 'Rm Category Not Defined'
				goto ErrHndlr
			end
			(Bef 2.12) ***/
			/* JewelSmith.2 RmCtg D added */
			select @row= count('x') from inserted ins where PMCD not in ('C','D') 
				/* not exists(Select 'x' from Param Pm where Pm.PTyp= 'RMCTG' and Pm.PMCd= ins.PValue) */
			if @row> 0 
			begin
				select @errmsg= 'Rm Category Should Be ''C'' Or ''D'' '
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/* JewelSmith.2 RmSctg must be removed, it should be blank */
			/****** PDesc225 = Blank ******/			
			select @row= count('x') from inserted where PDesc225 <> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
						
/*
			select @row= count('x') from inserted ins  where 
			 Exists	(Select 'x' From Param p where P.PTyp='RMSCTG'and P.PMCd=ins.PMCd and    
						 and (','+P.PSCd+',') Like ('%,'+p.PDesc225+',%') 
				 Group By p.PMCd,P.PSCd Having Count('x') > 1)
			if @row> 0 
			begin
				select @errmsg= 'Rm  Sub Category Cannot be enter more than ones'
				goto ErrHndlr
			end
*/

			/****** PValue1 = valid PSCd('RMSCTG') ******/
			select @row= count('x') from inserted ins where ins.PValue1<>'' 
			if @row> 0 
			begin
				select @errmsg= 'Rm Sub Category Should be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End
/******************************************Geeta*********************************/


		if @PTyp =  'PCT'
		Begin 
			/****** Cannot Add Records Of Ptyp= 'PCT' *****/
			if @mode in ('A') 
			begin
				select @errmsg= 'Cannot Add Records Of Param Type ''PCT'' '
				goto ErrHndlr
			end			
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank ******/
			select @row= count('x') from inserted where PDesc225= ''
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Not Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

    /***** Manali 3.9.0 - Output Folde For Report *****/     
		if @PTyp =  'OUTPUT'
		Begin 
			/****** Cannot Add Records Of Ptyp= 'PCT' *****/
			if @mode in ('A') 
			begin
				select @errmsg= 'Cannot Add Records Of Param Type ''OUTPUT'' '
				goto ErrHndlr
			end			
			/****** PmCd = OUTPUT ******/
			select @row= count('x') from inserted where PMCd<> 'OUTPUT'
			if @row>0 
			begin
				select @errmsg= 'Main Code Should Not Be Blank ''OUTPUT'' '
				goto ErrHndlr
			end
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank ******/
			select @row= count('x') from inserted where PDesc225= ''
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Not Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End

		if @PTyp = 'PLNDAYS'
		Begin 
			/****** Only One Record Of 'PLNDAYS' is allowed ******/
			select @row= count('x') from Param where PTyp= 'PLNDAYS' 
			if @row> 1
			begin
				select @errmsg= 'Only One Record Is Allowed For Ptyp= ''PLNDAYS'' '
				goto ErrHndlr
			end
			/****** Only One Record Of 'PLNDAYS' is allowed ******/
			select @row= count('x') from Param where PTyp= 'PLNDAYS' 
					and PMCd Not In ('PLNDAYS')
			if @row> 0
			begin
				select @errmsg= 'PMCd Should Be PLNDAYS'
				goto ErrHndlr
			end
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum > 0 ******/
			select @row= count('x') from inserted where PNum<= 0
			if @row> 0 
			begin
				select @errmsg= 'Number Of Days For Which Plan Has To Be Generated Should Be > Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = '' ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'PRDCTG' 
		Begin 
			/****** PSCd = Blank ******/			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank or valid Production Sequence ******/
/*			select @row=count('x') from inserted ins where ins.PValue<> '' and 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'PRDSEQ' and Prm.PMCd= ins.PValue)
			if @row>0 
			begin	
				select @errmsg= 'Production Sequence Not Defined'
				goto ErrHndlr
			end
*/
			/****** PValue = Blank ******/
			select @row=count('x') from inserted ins where ins.PValue<> ''
			if @row>0 
			begin	
				select @errmsg= 'PValue Should be Blank'
				goto ErrHndlr
			end

			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
	/*	Manoj 2.10.0	*/
			/****** PValidYN = Blank ******/
/*			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
*/
			select @row= count('x') from inserted ins where
				Not Exists (Select 'x' From Param Where PTYp = 'YN' And PMCD = ins.PValidYN)
			if @row>0
			begin
				select @errmsg= 'PValidYN Not Defined'
				goto ErrHndlr
			end
	/*	Manoj 2.10.0	*/
/*  #########################################  Manoj  #########################################  */
		End


		if @PTyp= 'PRDSTG' 
		Begin 
			if @mode= 'A' 
			begin 
				Select @ErrMsg= 'Cannot Add A New Production Stage'
				GoTo ErrHndlr
			end 

			select @Row= Count('x') from deleted del where @mode= 'C' and 
		    		not exists (Select 'x' from inserted ins where 
				ins.PTyp= del.PTyp and ins.PMCd= del.PmCd and ins.PSCd= del.PSCd 
				and Str(ins.PNum, 16, 0)= Str(del.PNum, 16, 0))
  			If @Row > 0
		  	begin 
				Select @ErrMsg= 'Cannot Change Priority Of Production Stage'
				GoTo ErrHndlr
			end
			/****** PSCd = Blank ******/			select @row= count('x') from inserted where PSCd<> ''
			if @row> 0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end

			/*** (Bef 2.14 Next) 
			****** PValue = Blank ******
			select @row= count('x') from inserted ins where ins.PValue<> ''
			if @row> 0 
			begin	
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			(Bef 2.14 Next) ***/

			/*** (Jen 2.14 Next) ***/
			/****** PValue <> Blank ******/
			select @row= count('x') from inserted ins where ins.PValue= ''
			if @row> 0 
			begin	
				select @errmsg= 'Caption Cannot Be Blank'
				goto ErrHndlr
			end
			/*** (Jen 2.14 Next) ***/

			/****** PNum >= 1 and PNum <= 5 ******/
			select @row= count('x') from inserted where Round(PNum, 0) < 1 or Round(PNum, 0) > 5 
			if @row> 0 
			begin
				select @errmsg= 'Priority Should Be Between 1 To 5'
				goto ErrHndlr
			end
			/****** PNum Should be Distinct ******/
			select @row= count('x') from inserted ins where 
				(Select count('x') from Param P1 where P1.PTyp= 'PRDSTG' 
					and Str(P1.PNum, 1, 0)= Str(ins.PNum, 1, 0))> 1
			if @row> 0 
			begin
				select @errmsg= 'Priority Should Be Unique For Each Prd Stage'
				goto ErrHndlr
			end
			/****** PNum shound be an integer ******/
			select @row= count('x') from inserted where Round(PNum, 0) <> Round(PNum, 4)  
			if @row> 0 
			begin
				select @errmsg= 'Priority Should Be An Integer'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end

			/*** (Bef 2.14 Next) 
			****** PValue1 = Blank ******
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			(Bef 2.14 Next) ***/

			/*** (Jen 2.14 Next) ***/
			/****** PValue1 = Blank or 'W-H' ******/
			select @row= count('x') from inserted ins where PValue1<> '' 
				and not exists (select 'x' from Param Prm where Prm.PTyp= 'W-H' and Prm.PMCd= ins.PValue1) 
			if @row> 0 
			begin
				select @errmsg= 'PValue1 Should Be Blank Or ''W'' Or ''H'' For Associating The WaxSet Or HandSet Diamond Availability Dates'
				goto ErrHndlr
			end

			/****** Only One 'W' Record allowed ******/
			select @row= count('x') from Param where PTyp= 'PRDSTG' and PValue1= 'W' 
			if @row> 1 
			begin
				select @errmsg= 'Only One ''W'' (WaxSet) Production Stage Allowed'
				goto ErrHndlr
			end

			/****** Only One 'H' Record allowed ******/
			select @row= count('x') from Param where PTyp= 'PRDSTG' and PValue1= 'H' 
			if @row> 1 
			begin
				select @errmsg= 'Only One ''H'' (HandSet) Production Stage Allowed'
				goto ErrHndlr
			end
			/*** (Jen 2.14 Next) ***/

			/****** PNum1 >= 0 and PNum1 <= 99 ******/
			select @row= count('x') from inserted where Round(PNum, 0) < 0 or Round(PNum, 0) > 99 
			if @row> 0 
			begin
				select @errmsg= 'Gap Days Should Be Between 0 To 99'
				goto ErrHndlr
			end
			/****** PNum1 shound be an integer ******/
			select @row= count('x') from inserted where Round(PNum, 0) <> Round(PNum, 4)  
			if @row> 0 
			begin
				select @errmsg= 'Gap Days Should Be An Integer'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row=count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'PTYP' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = 'Y' or 'N' ******/
			select @row=count('x') from inserted ins where 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'SYSDFN' and Prm.PMCd= ins.PValue)
			if @row>0 
			begin
				select @errmsg= 'Enter Type Of Parameter (Y/N/M)'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */

		End

		if @PTyp = 'RMCTG'
		Begin 
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank or Valid Raw Material Code of the same category ******/
			select @row=count('x') from inserted ins where PValue<>'' and 
				not exists (select 'x' from RmMst where RmCtg= PMCd and RmCd= PValue and RmZ= 'N')
			if @row>0 
			begin
				select @errmsg= 'Rm Code Used For Calculating Pure Wt Is Invalid'
				goto ErrHndlr
			end

			/****** Valid Raw Material Code with Q/W as 'W' ******/
			select @row=count('x') from inserted ins where PValue<>'' and 
				not exists (select 'x' from RmMst where RmCtg= PMCd and RmCd= PValue and RmQw = 'W' And RmQwCst = 'W')
			if @row>0 
			begin
				select @errmsg= 'The Pure RM Code Specified should always be defined as Weight Wise'
				goto ErrHndlr
			end

			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'RMGRP' 
		Begin 
			/****** PMCd = A report menu option like 'DSGCRD', 'ORDAVG', 'ORDDET' ******/			
			/****** PSCd = Not Blank ******/
			select @row=count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Rm Group Option'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
			goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'RMSCTG' 
		Begin 
			/****** PMCd = valid Raw Material Ctg ******/
			select @row= count('x') from inserted ins where 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'RMCTG' 
					and Prm.PMCd= ins.PMCd)
			if @row> 0 
			begin
				select @errmsg= 'Raw Material Category Not Defined'
				goto ErrHndlr
			end
			/****** PSCd = Not Blank ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row> 0 
			begin
				select @errmsg= 'Raw Material Sub Category Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue = valid Rate By Option ******/
			select @row= count('x') from inserted ins where 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'RATEBY' 
						and Prm.PMCd= ins.PValue)
			if @row> 0 
			begin
				select @errmsg= 'Rate By Option Not Defined'
				goto ErrHndlr
			end
			/****** PValue cannot be 'P' when rec exists in RmRt(RrTcTyp= 'PTR', RrCtg= PMCd, RrSCtg= PSCd)  ******/
			select @row= count('x') from inserted ins where PValue= 'P' and 
				exists (select 'x' from RmRt where RrTcTyp= 'PTR' and RrCtg= PMCd 
				and RrSCtg= PSCd and RrCmCtg= 'C' and RrSTWGrp= '' and RrCd= '' 
				and RrDmCtg= '' and RrLsCtg= '' and  RrLabMCd= '' and RrFrLn= 0 
				and RrToLn= 0) 
			if @row> 0 
			begin
				select @errmsg= 'RateBy Option Cannot Be Changed To ''P'' When Rec Exists In RmRt Entry Of Type ''PTR'' '
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum Value Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = 'Y' or 'N' ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'YN' 
									and Prm.PMCd= ins.PDesc225)
			if @row> 0 
			begin
				select @errmsg= 'Accept Stock By Size In Transaction (Y/N)'
				goto ErrHndlr
			end
			/****** PValue1 = 'Y' or 'N' ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'YN' 
									and Prm.PMCd= ins.PValue1)
			if @row> 0 
			begin
				select @errmsg= 'Accept Stock By Rate In Transaction (Y/N)'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where pdESC225 = 'N' AND PNum1 <> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = 'Y' or 'N' ******/
			select @row= count('x') from inserted ins where 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'YN' 
					    and Prm.PMCd= ins.PValue2)
			if @row> 0 
			begin
				select @errmsg= 'Accept Pointer Field In Design Rm (Y/N)'
				goto ErrHndlr
			end
			/****** PValue2 = 'N' when HPtrYN of Head is 'N' ******/
			select @row= count('x') from inserted ins where PValue2 = 'Y' and 
				(exists (select 'x' from Head where HCd= 'ZSELF' and HPtrYN= 'N')) --or 
				-- 6.1 NEW PMCd not in ('C', 'D'))
			if @row> 0 
			begin
				select @errmsg= ' ''Dsg Ptr (Y/N)'' Field Should Be ''N'' '
				goto ErrHndlr
			end
			/****** PValue3 = 'Y' or 'N' ******/
			select @row= count('x') from inserted ins where 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'YN' 
				and Prm.PMCd= ins.PValue3)
			if @row> 0 
			begin
				select @errmsg= 'Accept Stock By Pointer In Transaction (Y/N)'
				goto ErrHndlr
			end
			/****** PDesc225 and PValue3 Cannot Be 'Y' together ******/
			select @row= count('x') from inserted ins where PDesc225 = 'Y' and PValue3= 'Y' 
			if @row> 0 
			begin
				select @errmsg= 'Stock Cannot Be Maintained Both Ptr Wise And Sz Wise In Transaction (Y/N)'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'RMCOL' 
		Begin 
			/****** PMCd = valid Raw Material Ctg ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'RMCTG' 
								and Prm.PMCd= ins.PMCd)
			if @row>0 
			begin
				select @errmsg= 'Raw Material Category Not Defined'
				goto ErrHndlr
			end
			/****** PSCd = Not Blank ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'Raw Material Colour Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted ins where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Value Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted ins where PDesc225<> '' 
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted ins where PValue1<> ''
			if @row>0 
			begin
			select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
	/*	Manoj 2.10.0	*/
			/****** PValidYN = Blank ******/
/*			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
*/
			select @row= count('x') from inserted ins where
				Not Exists (Select 'x' From Param Where PTYp = 'YN' And PMCD = ins.PValidYN)
			if @row>0
			begin
				select @errmsg= 'PValidYN Not Defined'
				goto ErrHndlr
			end
	/*	Manoj 2.10.0	*/
/*  #########################################  Manoj #########################################  */
		End


		/*** (Jen 2.14 Next) ***/
		if @PTyp = 'RTC'
		Begin 
			/****** Only One Record Of Each PMCD ('LOCKDT' & 'FUTUREDT')
				 is allowed ******/
			select @row= count('x') from Param where PTyp= 'RTC' 
					and PMCd Not In ('RTC') 
			if @row> 0
			begin
				select @errmsg= 'PMCd Should Be ''RTC'' '
				goto ErrHndlr
			end			
			/****** PSCd >= 0 ******/
			--MW.105 - RTC Sr. should not be empty			
			select @row= count('x') from inserted where convert (int, PSCd)<= 0
			if @row> 0 
			begin
				select @errmsg= 'Sr. Should Be > Zero'
				goto ErrHndlr
			end
			
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum >= 0 ******/
			select @row= count('x') from inserted where convert (int, PNum) < 0 
			if @row> 0 
			begin
				select @errmsg= 'PNum Should Be >= Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 >= 0 ******/
			select @row= count('x') from inserted where convert (int, PNum1)< 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be >= Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row> 0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		/*** (Jen 2.14 Next) ***/


		/****** Sachin - Multi Server  ******/
		if @PTyp = 'REPCTR'
		Begin 
			/****** Only One Record Of Each PMCD ('LOCKDT' & 'FUTUREDT')
				 is allowed ******/
			select @row= count('x') from Param where PTyp= 'REPCTR' 
					and PMCd Not In ('REPCTR') 
			if @row> 0
			begin
				select @errmsg= 'PMCd Should Be ''REPCTR'' '
				goto ErrHndlr
			end
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row> 0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum >= 0 ******/
			select @row= count('x') from inserted where convert (int, PNum) < 0 
			if @row> 0 
			begin
				select @errmsg= 'PNum Should Be >= Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 >= 0 ******/
			select @row= count('x') from inserted where convert (int, PNum1)< 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be >= Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row> 0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		/****** Sachin - Multi Server ******/

		if @PTyp= 'SEOPWD' 
		Begin 
			/****** Only One Record Of Type 'VER' is allowed ******/
			select @row= count('x') from Param where PTyp= 'SEOPWD' 
			if @row> 1 
			begin
				select @errmsg= 'Only One SEO Passowrd Can Be Entered'
				goto ErrHndlr
			end
			/****** PMCd = 'SEOPWD' ******/
			select @row= count('x') from inserted ins where ins.PMCd<> 'SEOPWD'
			if @row>0 
			begin
				select @errmsg= 'Main Code Should Be ''SEOPWD'' '
				goto ErrHndlr
			end
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Value Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225= Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end			/****** PValue1 = Can Be Blank ******/
			/* select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			*/

			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SRTBAG' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid fields of the Bag table ******/	
			select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
			select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SRTDLD' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid fields of the Bag table ******/	
			select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
			select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SRTINBAG' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid fields of the InvFgd table ******/	
			/*select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end*/
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SRTPDC' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid or not ******/	
			select @row= count('x') from inserted where PDesc225= ''
			if @row> 0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
			select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SRTPRHLP' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid fields of the TxndZ table ******/	
			select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  ######################################### Manoj  #########################################  */
		End

		if @PTyp= 'SRTLR' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid fields of the LabRt table ******/	
			select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SRTPP' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid fields of the PrdPts table ******/	
			select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SRTRM' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid fields of the RmMst table ******/	
			select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SRTRR' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end			/****** PDesc225 = Not Blank (it is a valid sequence of fields on which the sorting has to be done) ******/
			/****** It is difficult to check if the individual fields in the sequence are valid fields of the RmRt table ******/	
			select @row=count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Enter Sort On Fields'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'SUBPRC' 
		Begin 
			/* Check Below Shifted to NR Trigger in version 2.11.0 */
			/****** PMCd = valid Process Loc ******/
			/*select @row= count('x') from inserted where not exists (select 'x' from Loc where LocTyp= 'P' and LocCd= PMCd)
			if @row>0 
			begin
				select @errmsg= 'Process Location Not Defined'
				goto ErrHndlr
			end*/
			/****** PSCd = Not Blank ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'Sub Process Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin	
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
			goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

/*************Bhavna valids for RWRK *********/
		if @PTyp= 'RWRK' 
		Begin 
			/****** PSCd = Not Blank ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'ReWork Code Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin	
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
			goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end

			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end

		End

/*****************Bhavna Valods for RWRK**********/

		/*** (Jen 2.13) (07/11/06) ***/
		if @PTyp= 'SYSOPT' 
		Begin 
			/****** Cannot Add Recs Of This PTyp ******/
			if @mode in ('A') 
			begin
				select @errmsg= 'Cannot Add Records Of SYSOPT Param Type' 
				goto ErrHndlr
			end

			/****** Cannot Change Description ******/
	  		select @Row= Count('x') from deleted del where 
	    		not exists (Select 'x' from inserted ins where ins.PTyp= del.PTyp and ins.PMCd= del.PmCd and ins.PSCd= del.PSCd and ins.PDesc= del.PDesc)
	  		If @Row > 0 
		  	begin 
				Select @ErrMsg= 'Cannot Change Description'
				GoTo ErrHndlr
			end

			/****** PSCd = Not Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'PSCd Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = valid Param(PTyp= 'YN') ******/
			select @row= count('x') from inserted ins where 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'YN' 
				and Prm.PMCd= ins.PValue)
			if @row> 0 
			begin
				select @errmsg= 'Option Should Be ''Y'' or ''N'' '
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin	
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
			goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end

			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		/*** (Jen 2.13) (07/11/06) ***/

	
		if @PTyp= 'TC' 
		Begin 
			/****** Cannot Add Records Of Ptyp= 'TC' *****/
			if @mode in ('A') 
			begin
				select @errmsg= 'Cannot Add Records Of Param Type ''TC'' '
				goto ErrHndlr
			end	
			/****** Can Change Only The Default 'CHR' ******/
	  		select @Row=Count('x') from deleted del where 
    			not exists (Select 'x' from inserted ins where ins.PTyp= del.PTyp 
					and ins.PMCd= del.PmCd and ins.PSCd= del.PSCd and 
					ins.PDesc= del.PDesc and ins.PValue= del.PValue and
					/* ins.PDesc225= del.PDesc225 and  */
					Round(ins.PNum, 4)= Round(del.PNum, 4) and 
					ins.PValue1= del.PValue1 and 
					Round(ins.PNum1, 4)= Round(del.PNum1, 4) and 
					ins.PValue2= del.PValue2 and 
					ins.PValue3= del.PValue3)
  			If @Row > 0
		  	begin 
				Select @ErrMsg= 'Cannot Change Any Other Value Other Than Default CHR'
				GoTo ErrHndlr
			end
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = valid Tc Type ******/
			select @row= count('x') from inserted ins where 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'TCTYP' 
				and Prm.PMCd= ins.PValue)
			if @row>0 
			begin
				select @errmsg= 'Tc Type Not Defined'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted ins where ins.PDesc225<> '' 
			if @row>0 
			begin
				select @errmsg= 'PDesc225 should be blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'TCTYP' 
		Begin 
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank or valid Tc of the given Tc Type ******/
			select @row=count('x') from inserted ins where ins.PValue<> '' and 
				not exists (select 'x' from Param Prm where Prm.PTyp= 'TC' 
					and Prm.PMCd= ins.PValue and Prm.PValue= ins.PMCd)
				if @row>0 
			begin
				select @errmsg= 'Tc Not Defined'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row=count('x') from inserted where PNum<> 0
			if @row>0 
			begin	
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row=count('x') from inserted where PDesc225<> '' 
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = valid Entry Type ******/
			select @row= count('x') from inserted where PValue1 not in('Om', 'TxnBag', 'TxnRm', 'Tz', 'FgBag', 'FgRm', 'In', 'FgMvm')
			if @row> 0 
			begin
				select @errmsg= 'Enter ''Om'', ''TxnBag'', ''TxnRm'', ''Tz'', ''FgBag'', ''FgRm'' Or  ''In'', ''FgMvm'' '
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp = 'TRFDIR'
		Begin 
			/****** Only One Record Of Each PMCD ('LOCKDT' & 'FUTUREDT')
				 is allowed ******/
			select @row= count('x') from Param where PTyp= 'TRFDIR' 
					and PMCd Not In ('TRFDIR')
			if @row> 0
			begin
				select @errmsg= 'PMCd Should Be ''TRFDIR'' '
				goto ErrHndlr
			end
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = a valid path ******/
			select @row= count('x') from inserted where PDesc225= ''
			if @row> 0 
			begin
				select @errmsg= 'Default Transfer Directory Path Should Be Entered'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'USR' 
		Begin 
			/* *** Jay 2.14 *** (Pvalue <> '' when UaSeoYn = 'Y'  )*/
			select @row= count('x') from inserted ins where PValue= '' and 
				Exists (Select 'x' From UsrAccess where UaUsrCd= ins.PMCd and UaSeoYn= 'Y')
			if @row>0
			begin
				select @errmsg= 'Password Can Not Be Blank When User Has SEO Rights'
				goto ErrHndlr
			end
			
			/****** PSCd = Blank ******/
			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/*** No Validation For PValue (password is entered in this field) ***/

/* ###			****** PDesc225 = Blank or Valid Sequence of Process Locations ******
			select @row=count('x') from inserted where PDesc225<> '' and 
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank Or A Valid Loc Seq'
				goto ErrHndlr
			end
			****** PValue1 = Blank or Valid Sequence of Non Process Locations ******
			select @row=count('x') from inserted where PValue1<> '' and 
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
*/
			/*** Shilpa Emr 2.04 Disable the Upper Code for PNUM where Ptyp = 'USR'  ****/
			/****** PNum > 1 and Pnum <= 9 ******/
			select @row= count('x') from inserted where Round(PNum, 0)< 1 or Round(Pnum, 0)> 9 
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Between 1 To 9'
				goto ErrHndlr
			end

			/****** PNum shound be an integer ******/
			select @row= count('x') from inserted where Round(PNum, 0) <> Round(PNum, 4) 
			if @row> 0 
			begin
				select @errmsg= 'Priority Should Be An Integer'
				goto ErrHndlr
			end
			/*** Shilpa Emr 2.04 ****/

			/* Jenny Color (02/04/05) */
			/****** If PValue2 <> Blank then valid CsCd from ClrSet ******/
			select @row= count('x') from inserted where PValue2<> '' 
				and Not Exists (Select 'x' From ClrSet Where CsCd= PValue2) 
			if @row> 0 
			begin
				select @errmsg= 'Color Code Not Defined (It Can Be Blank)'
				goto ErrHndlr
			end
			/* Jenny Color (02/04/05) */

			/*** (Jen 3.01) ***/ 
			/****** PValue3 = valid User Group ******/
			/****** Manali 3.7.1 - 'JMIC' added ******/
			/* */ 
			select @row= count('x') from inserted ins where ins.PMCd <>'MIC'
				and Not Exists (Select 'x' From Param P Where P.PTyp= 'USRGRP' and P.PMCd= ins.PValue3) 
			if @row> 0 
			begin
				select @errmsg= 'User Group Not Defined'
				goto ErrHndlr
			end

			/****** Manali 3.7.1 - 'JMIC' added ******/
			select @row= count('x') from inserted ins where ins.PMCd ='MIC'
				and ins.PValue3<> '' 
			if @row> 0 
			begin
				select @errmsg= 'User Group Should Be Blank' 
				goto ErrHndlr
			end
			/*** (Jen 3.01) ***/

/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
	/*	Manoj 2.10.0	*/
			/****** PValidYN = Blank ******/
	/*			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
	*/
			select @row= count('x') from inserted ins where
				Not Exists (Select 'x' From Param Where PTYp = 'YN' And PMCD = ins.PValidYN)
			if @row>0
			begin
				select @errmsg= 'PValidYN Not Defined'
				goto ErrHndlr
			end 

	/* Zubin 212 */
			/****** (Jen 2.14 Next Patch 1) added 'ADM' ******/
			/****** Manali 3.7.1 - 'JMIC' added ******/
			select @row= count('x') from inserted where PMCd= 'MIC' and PValidYN= 'N' 
			if @row>0
			begin
				select @errmsg= 'PValidYN Cannot Be ''N'' For Users ''MIC'''
				goto ErrHndlr
			end 
	/* Zubin 212 */

			select @row= count('x') from inserted ins where PMCd= 'MIC' and PNum1<> 99 
			if @row> 0 and (Update(PDesc) or Update(PDesc225) or 
				Update(PNum) or Update(PValue1) or Update(PValue2) or Update(PValue3) or Update(PValidYn)) 
			begin
				select @errmsg= 'Cannot Update ''MIC'' User' 
				goto ErrHndlr
			end 
				
			select @row= count('x') from inserted ins where PMCd= 'MIC' and PNum1<>1
			if @row>0
			begin				
				 Update Param set PNum1=1 where PTyp= 'USR' and PMCd= 'MIC' 
			end 			
			/* Zubin 212 (05/09/06) */ 

			/* ****** Sachin 3.2.0 - Checking PNum1 Value ****** */
			select @row= count('x') from inserted ins where PMCd<> 'MIC' and
				Not Exists (Select 'x' From Param Prm where Prm.PTyp='PTNTYP' And Prm.PNum1=ins.PNum1)
			if @row>0
			begin
				select @errmsg= 'Partition Serial Not Defined ' 
				goto ErrHndlr
			end
			/* ****** Sachin 3.2.0 - Checking PNum1 Value ****** */
	
			/****** Cannot Change Partition Key from 2 to 1 is the user is already logged on ******/
			select @row= count('x') from inserted ins where ins.PNum1=1
				and Exists (Select 'x' From Deleted del Where ins.PTyp=del.PTyp and ins.PMCd=del.PMCd and ins.PSCd=del.PSCd and ins.PNum1<>del.PNum1)
				and Exists (Select 'x' From UsrLogin Where ins.PMCd=UlUsrCd ) 
			if @row> 0 
			begin
				select @errmsg= 'Cannot Change the Partition Serial as User Already Logged On'
				goto ErrHndlr
			end


	/*	Manoj 2.10.0	*/
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'VACTG' 
		Begin 
			/****** PSCd = Blank ******/			select @row=count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum >= 0 ******/
			select @row= count('x') from inserted where PNum< 0
			if @row> 0 
			begin
				select @errmsg= 'Enter Proper Value Addition Gold Loss % '
				goto ErrHndlr
			end
		
			/****** PValue1 = Blank ******/
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			/***Emr208*******Geeta******/
		/*	select @row=count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
		*/
			/****** PValue2 = Blank ******/
			select @row=count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp= 'VER' 
		Begin 
			/*** (Jen 3.01) ***/
			if @mode in ('C') 
			Begin
				/*** Note that the PValidYN field is editable. The user will not be able to log in to the system 
				if pValidYN= 'N'. Also this field will remain invisible to the user ***/
  				select @Row= Count('x') from deleted del where
    				not exists (Select 'x' from inserted ins where ins.PTyp= del.PTyp and ins.PMCd= del.PmCd 
								and ins.PSCd= del.PSCd and ins.PDesc= del.PDesc and ins.PValue= del.PValue 
								and ins.PValue1= del.PValue1 and ins.PValue2= del.PValue2 
								and ins.PDesc225= del.PDesc225)  
  				If @Row > 0
	  			begin 
					Select @ErrMsg='Cannot Change ''VER'' Details'
					GoTo ErrHndlr
				end
			End 
			/*** (Jen 3.01) ***/

			/****** Only One Record Of Type 'VER' is allowed ******/
			select @row= count('x') from Param where PTyp= 'VER' 
			if @row> 1 
			begin
				select @errmsg= 'Only One Version Record Can Be Entered'
				goto ErrHndlr
			end
			/****** PMCd = 'DATVER' ******/
			select @row= count('x') from inserted ins where ins.PMCd<> 'DATVER'
			if @row>0 
			begin
				select @errmsg= 'Main Code Should Be ''DATVER'' '
				goto ErrHndlr
			end
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			/* select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			*/
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Value Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225<> Blank ******/
			select @row= count('x') from inserted where PDesc225= ''
			if @row>0 
			begin
				select @errmsg= 'Database Version Should Be Entered'
				goto ErrHndlr
			end			/****** PValue1 = Blank ******/
			/* select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			*/
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			/*select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			*/
			/****** PValue3 = Blank ******/
			/* select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			*/
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			/*** (Bef 3.01)
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
			(Bef 3.01) ***/
/*  #########################################  Manoj  #########################################  */
		End
	/**********************Geeta*********************Emr206*******************/
	if @PTyp= 'TRCHR' 
		Begin 
			/****** PMCd= Not Blank ******/
			select @row= count('x') from inserted where PMCd= ''
			if @row> 0 
			begin
				select @errmsg= 'Main Code Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row> 0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue= Blank ******/
			select @row= count('x') from inserted where PValue<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** Round(PNum, 3)<> 0 ******/
			select @row= count('x') from inserted where Round(PNum, 3)<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225= Blank ******/
			select @row= count('x') from inserted where PDesc225<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1= Blank ******/		
			select @row=count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1= 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row> 0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row> 0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end

		End

	if @PTyp = 'WAXSPGR'
		Begin 
		/****** Only One Record Of  PMCD='WAXSPGR is allowed ******/
			select @row= count('x') from Param where PTyp= 'WAXSPGR' 
					and PMCd <> 'WAXSPGR'
			if @row> 0
			begin
				select @errmsg= 'PMCd Should Be Only ''WAXSPGR''(Only One Record can be Entered) '
				goto ErrHndlr
			end
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/

		/*	select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end         */

			/****** PDesc225 = a valid date ******/
			select @row= count('x') from inserted where PDesc225<>''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end

			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end

		End


		/* **** Manali 3.5.0 REPCOl  **** */
		If @PTyp='REPCOL' 
		Begin
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end

			/****** PDes<> Blank ******/
			select @row= count('x') from inserted where PDesc=''
			if @row>0
			begin
				select @errmsg= 'Descripition Cannot be Blank'
				goto ErrHndlr
			end
	
			/****** PDesc225 <> Blank ******/
			select @row= count('x') from inserted where PValue = ''
			if @row>0
			begin
				select @errmsg= 'Colour Code Cannot Be Blank'
				goto ErrHndlr
			end

		End
		/* **** Manali 3.5.0 WIWCOL  **** */


		/* **** Manali 3.5.0 WIWCOL  **** */
		If @PTyp='WIWCOL' 
		Begin
			Select @row= count('x') from Inserted Ins 
				where Not Exists( Select 'x' From Param Pm where Pm.PTyp='REPCOL' and Pm.PMCd= ins.PMCD)
			If @row > 0 
			Begin	
				select @errmsg= 'Colour not defined'
				goto ErrHndlr
			End
			
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end

			/****** PValue < PValue1 ******/
			select @row= count('x') from inserted where
				PNum > PNum1
			if @row>0
			begin
				select @errmsg= 'From Days Should Be less Than To Days'
				goto ErrHndlr
			end
			/* **** Manali - 3.6.0 - Following Check commented as negative allowed in range **/
/*			Select @row= count('x') from inserted where PNum1 = 0 
			If @Row > 0 
			Begin
				Select @errmsg= 'To Days cannot be zero'
				goto Errhndlr
			End
*/
			Select @row = count('x') from inserted ins 
			where 
				 (Select count('x') from Param Pm where Pm.PTyp='WIWCOL' and		
				((Ins.PNum between Pm.PNum and Pm.PNum1) Or (Ins.PNum1 between Pm.PNum and Pm.PNum1))) > 1
				
			If @row > 0 
			Begin	
				Select @errmsg= 'Range already defined'
				goto Errhndlr
			End

			
		End
		/* **** Manali 3.5.0 WIWCOL  **** */

		/**********************Geeta*********************Emr206*******************/
		/********* ZUBIN - EMR206, 10th Dec 2003 START @PTyp= 'PRTCD' *********/
		if @PTyp= 'PRTCD'
		Begin
			/****** PSCd= Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end

			/******* PValue/1/2/3 should be Numeric *******/
			select @row = count('x') from inserted Ins where
				IsNumeric(PValue) <> 1 Or IsNumeric(PValue1) <> 1 Or
				IsNumeric(PValue2) <> 1 Or IsNumeric(PValue3) <> 1
			if @row > 0
			begin
				select @errmsg= 'Dia%% Fr/Dia%% To/Del Days Fr/Del Days To Should Be Numeric'
				goto ErrHndlr
			end

			/****** PDesc225 = Y/N ******/
			select @row = count('x') from inserted Ins where
				NOT EXISTS (Select 'x' from Param where PTyp = 'YN' and PMCd = Ins.PDesc225)
			if @row > 0
			begin
				select @errmsg= 'ReAllot Priority Should Be Y/N'
				goto ErrHndlr
			end
		/**************************************Geeta*************************************/
		/********************* PValue/PValue1/PValue2/PValue3 cann't be blank when PDesc225 is 'N' */
			/****** PValue/PValue1/PValue2/PValue3 blank when PDesc225 is 'N' ******/
		/*	select @row= count('x') from inserted where PDesc225='N' and
				(Cast(PValue as float) <> 0 or Cast(PValue1 as float) <> 0 or
				Cast(PValue2 as float) <> 0 or Cast(PValue3 as float) <> '')
			if @row>0
			begin
				select @errmsg= 'Dia%% Fr/Dia%% To/Del Days Fr/Del Days To Should Be Blank When ReAllot Priority Is 'N''
				goto ErrHndlr
			end  */

			/**************** PValue and PValue1 cannot be < 0 ************************/
			select @row= count('x') from inserted where
				Cast(PValue as float) NOT BETWEEN 0 and 100 or
				Cast(PValue1 as float) NOT BETWEEN 0 and 100
			if @row > 0
			begin
				select @errmsg= 'Dia%% Fr and Dia%% To Should Be Between 0 and 100'
				goto ErrHndlr
			end

			/****** PValue < PValue1 ******/
			select @row= count('x') from inserted where
				Cast(PValue as float) > Cast(PValue1 as float)
			if @row>0
			begin
				select @errmsg= 'Dia%% Fr Should Be <= Dia%% To'
				goto ErrHndlr
			end

			/****** PValue2 < PValue3 ******/
			select @row= count('x') from inserted where
				Cast(PValue2 as float) > Cast(PValue3 as float)
			if @row>0
			begin
				select @errmsg= 'Del Days Fr Should Be <= Del Days To'
				goto ErrHndlr
			end

			/*** Valid PValue-PValue1 Range ***/
			select @row= count('x') from inserted ins where  
				(Select count('x') from Param Pm where Pm.PTyp = ins.PTyp 
				and  ( (Cast(Pm.PValue as float) <= Cast(Ins.PValue as float) and
							Cast(Pm.PValue1 as float) >= Cast(Ins.PValue as float))
				or   (Cast(Pm.PValue as float) <= Cast(Ins.PValue1 as float) and
							Cast(Pm.PValue1 as float) >= Cast(Ins.PValue1 as float))
			     or   (Cast(Ins.PValue as float) <= Cast(Pm.PValue as float) and
							Cast(Ins.PValue1 as float) >= Cast(Pm.PValue1 as float))
			)) > 1
			if @row> 0
			Begin
				select @errmsg= 'Dia%% Range Already Defined'
				goto ErrHndlr
			end

			/*** Valid PValue2-PValue3 Range ***/
			select @row= count('x') from inserted ins where 
				(Select count('x') from Param Pm where Pm.PTyp = ins.PTyp 
		   		and  ( (Cast(Pm.PValue2 as float) <= Cast(Ins.PValue2 as float) and
							Cast(Pm.PValue3 as float) >= Cast(Ins.PValue2 as float)  )
				     or   (Cast(Pm.PValue2 as float) <= Cast(Ins.PValue3 as float) and
							Cast(Pm.PValue3 as float) >= Cast(Ins.PValue3 as float)  )
		       		     or   (Cast(Ins.PValue2 as float) <= Cast(Pm.PValue2 as float)   and
							Cast(Ins.PValue3 as float) >= Cast(Pm.PValue3 as float)  )
					 )) > 1
			if @row> 0
		 	Begin
		  		select @errmsg= 'Del Days Range Already Defined'
		  		goto ErrHndlr
		 	end

			/****** PValue-PValue1 or PValue2-PValue3 not blank when PDesc225 is 'Y'******/
		/*	select @row= count('x') from inserted where 
				(Cast(PValue as float) + Cast(PValue1 as float) <= 0) or
				((Cast(PValue2 as float) = 0 And Cast(PValue3 as float) = 0))
			if @row>0
			begin
				select @errmsg= 'Dia%% Fr and Dia%% To Or Del Days Fr and Del Days To Should Not Be Zero'
				goto ErrHndlr
			end  */

			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum <> 0
			if @row>0
			begin
				select @errmsg= 'PNum Value Should Be Zero'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
		End
		/********* ZUBIN - EMR206, 10th Dec 2003 END @PTyp= 'PRTCD' *********/


/****** (Jen 3.01) (added 'USRGRP') ******/
/* **** Zubin 211 ('BCHR' to be removed) **** */ 
/****** (Jen 2.13) 'RETCHR', 'SRTRTHLP' added ******/
/* **** Manali 3.6.0 - 19/09/09 - 'DSGEXC' added **** */
		if @PTyp in (/*'BCHR', */'BAGPCS', 'BHISTOPT', 'BRKMSG', 'CMCTG', 'DC', 
			     'DMTCTYP', 'DMUOM', 'DTTAG', 'FA', 'FNDOPT', 
			     'IMPTYP', 'INVAL', 'INVTOTAS', 'INWT', 'IR', 'KT', 'LABMCD', 
			     'LABQW', 'LOCTYP', 'LSCTG', 'ORDVAL', 'PDCSFX', 'PLCHR', 
			     'PRDTYP', 'PURCPY', 'RATEBY', 'REGION', 'REJCD', 'RETCHR', 
			     'RMUOM', 'RNDOPT', 'RRTCTYP', 'SHPCDGR', 'SP', 'SRTRTHLP', 
			     'STWGRP', 'SUBSHP', 'USRGRP', 'VCHOPT', 'W-H', 'WH', 'YN', 'YY', 
			     'ZWTEQ','LOOKS','SETTYP', 'DSGEXC')
/* **** Zubin 211 **** */ 
		Begin 
			/* '*** Jay 2.14 *** */
			/**************** For Yy Field Check Numeric Value ************************/
			If @Mode= 'A' and @PTyp = 'YY'
			Begin
				Select @row= count('x') from inserted where  IsNumeric (PMCd) <> 1 
				If @row > 0
				Begin
					Select @errmsg= 'Yy Field Should Be Numeric'
					Goto ErrHndlr
				End
			End
			/* '*** Jay 2.14 *** */
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
--	Manoj
			/****** PValue = Blank ******/
/*			select @row= count('x') from inserted where PValue<> '' And PTyp <> 'REGION'
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
*/
			/****** PValue = Blank (Not For 'Region') ******/
			select @row= count('x') from inserted where PValue<> '' And PTyp <> 'REGION'
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
--	Manoj
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end

			/****** PDesc225 = Blank 4.1.4 ******/  
			select @row= count('x') from inserted where @PTyp<>'RETCHR' And PDesc225<>''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
/*  #########################################  Manoj  #########################################  */
		End

		if @PTyp in ('DMCOL', 'DMSZ','SALCTG','CELL','DTCYC','PRDLINE')
		Begin 
			
			/********Geeta****CRM 2.12
				**PMCD= Valid PMCD('DMCTG') and PSCD<>'' For PTYP='SALCTG'***/
		     If @PTyp ='SALCTG' 
			Begin
				/****** PMCd =Valid PMCD('DMCTG') ******/
				select @row= count('x') from inserted  ins where 
					Not Exists( Select 'x' From Param Pm where PTyp='DMCTG' and Pm.PMCd= ins.PMCD)
				if @row>0 
				begin
					select @errmsg= 'DmCtg Not Defined'
					goto ErrHndlr
				end

				/****** PSCd <>'' ******/
				select @row= count('x') from inserted where PSCd= ''
				if @row>0 
				begin
					select @errmsg= 'Sales Category Cannot be Blank'
					goto ErrHndlr
				end
			End
			/*******Geeta**************/
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> '' and PTyp<>'SALCTG'
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
--	Manoj
			/****** PValue = Blank ******/
			select @row= count('x') from inserted where PValue<> '' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
--	Manoj
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
	/*	Manoj 2.10.0	*/
			/****** PValidYN = Blank ******/
/*			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
*/
			select @row= count('x') from inserted ins where
				Not Exists (Select 'x' From Param Where PTYp = 'YN' And PMCD = ins.PValidYN)
			if @row>0
			begin
				select @errmsg= 'PValidYN Not Defined'
				goto ErrHndlr
			end
	/*	Manoj 2.10.0	*/
/*  #########################################  Manoj  #########################################  */
		End
/* *** Jay 2.13(CT) *** */
/* **** Zubin 211 **** */ 
		if @PTyp= ('BCHR')
		Begin
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/* PValue= 'YN' */ 
			select @row= count('x') from inserted ins where ins.PTyp= 'BCHR' and 
				Not Exists (select 'x' from Param Prm where Prm.PTyp= 'BTYP' and 
					Prm.PMCd= ins.PValue and Prm.PSCd= '')
			if @row> 0 
			begin 
				select @errmsg= 'Invalid Bag Type' 
				goto ErrHndlr 
			end 
			/* Cannot create a flute bag if hFluteBagYN= 'N' in Head table for CoCd= 'ZZZ' */ 
			/* (Bef 22/06/05) select @row= count('x') from inserted ins where PTyp= 'BCHR' and */
			select @row= count('x') from inserted ins where PTyp= 'BCHR' and PValue= 'F' and 
				Not Exists (select 'x' from Head where HCoCd= 'ZZZ' and hFluteBagYN= 'Y') 
			if @row>0 
			begin
				select @errmsg= 'Cannot Create Flute Bag Character As The Option For Flute Bag= ''N'' In Head Table' 
				goto ErrHndlr 
			end 
			/* Cannot Change Bag Character Type once a bag is created */ 
			if @mode = 'C'
			begin
				select @row= count('x') from inserted ins where ins.PTyp= 'BCHR' and 
					Not Exists (Select 'x' from deleted del where ins.PTyp= del.PTyp and 
						ins.PMCd= del.PMCd and ins.PSCd= del.PSCd and ins.PValue= del.PValue) and 
					Exists (Select 'x' from Bag where BChr= ins.PMCd) 
				if @row> 0 
				begin
					select @errmsg= 'Cannot Change Bag Option As Bag Chr In Use In Bag Table' 
					goto ErrHndlr 
				end 
			end

			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			select @row= count('x') from inserted where PDesc225<> '' and 
				PValue <> 'F' 
			if @row>0 
			begin
				select @errmsg= 'SubShp Should Be Enter For Only Flute Bags' 
				goto ErrHndlr
			end
/* **** Zubin 211 **** */
			/****** PValue1 = Blank for primary bags & Components COdes ******/
			select @row= count('x') from inserted where PValue1<> '' and 
				PValue <> 'F' 
			if @row>0 
			begin
				select @errmsg= 'Rm Category Should Be Enter For Only Flute Bags' 
				goto ErrHndlr
			end
				
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank for primary bags  and Component Bags ******/
			select @row= count('x') from inserted where PValue2<> '' and PValue<> 'F' 
			if @row>0 
			begin
				select @errmsg= 'Rm Sub Category Should Be Enter For Only Flute Bags'
				goto ErrHndlr
			end

	/* ***** Not required as multiple RmSubCtg can be selected ***** */ 
			/****** PValue2 = blank or Valid RmSCtg for Flute Bags ******/
/*			select @row= count('x') from inserted ins where PValue2<> '' and 
				PValue= 'Y' and 
				Not Exists (Select 'x' from Param Prm where 
					Prm.PTyp= 'RMSCTG' and Prm.PMCd= ins.PValue1 and 
					Prm.PSCd= ins.PValue2) 
			if @row>0 
			begin
				select @errmsg= 'Rm Sub Category Is Invalid'
				goto ErrHndlr
			end
*/
		
/* *** Jay 2.13(CT) *** */
			select @row= count('x') from inserted ins where PDesc225<> ''  
				 and 
				Not Exists (Select 'x' from Param Prm where 
					Prm.PTyp= 'SUBSHP' and Prm.PMCd= ins.PDesc225  and 
					Prm.PSCd= '') 
			if @row>0 
			begin
				select @errmsg= 'SubShp Is Invalid'
				goto ErrHndlr
			end
			select @row= count('x') from inserted ins where PTyp= 'BCHR' and PValue= 'C' and 
				Not Exists (select 'x' from Head where HCoCd= 'ZZZ' and hCompBagsYn= 'Y') 
			if @row>0 
			begin
				select @errmsg= 'Cannot Create Component Bags Character As The Option For Component Bags= ''N'' In Head Table' 
				goto ErrHndlr 
			end 

			Select @row= count('x') from Param P1 where PTyp = 'BTYP' and PMCd= 'C' and PSCd= '' 
				and convert(int,PNum) < IsNull((Select count('x') From Param P2 where P2.PTyp= 'BCHR' and P2.PValue= 'C'), 0) 
			if @row> 0 
			begin
				select @errmsg= 'Can Not Add More Then Specified Component Cherecter Codes'
				goto ErrHndlr
			end						

			/****** Cannot Change PValue if BChr is used in DsgMst table ******/
			select @row= count('x') from deleted del where del.PValue = 'C' and 
				exists (select 'x' from DsgMst where DmTcTyp = 'PM' and DmSz= Del.PMCd)
				and del.PValue <> (Select PValue from Inserted ins where del.PTyp= ins.PTyp and del.PMCd= ins.PMCd  and del.PSCd= ins.PSCd )
			if @row> 0 
			begin				select @errmsg= 'Cannot Change Bag Type as Used in Design Master Table'
				goto ErrHndlr
			end

			/****** Cannot Change PValue if BChr is used in DsgCT table ******/			
			select @row= count('x') from deleted del where del.PValue = 'C' and 
				exists (select 'x' from DsgCT where DctTcTyp= 'DM' and DctDmSz= '' and DctChr= del.PMCd)
				and del.PValue <> (Select PValue from Inserted ins where del.PTyp= ins.PTyp and del.PMCd= ins.PMCd  and del.PSCd= ins.PSCd )
			if @row> 0 
			begin				select @errmsg= 'Cannot Change Bag Type as Used In Design Component Table'
				goto ErrHndlr
			end

			/****** Cannot Change PValue if BChr is used in OrdCT table ******/			
			select @row= count('x') from deleted del where del.PValue = 'C' and 
				exists (select 'x' from OrdCT where OctCtChr= del.PMCd)
				and del.PValue <> (Select PValue from Inserted ins where del.PTyp= ins.PTyp and del.PMCd= ins.PMCd  and del.PSCd= ins.PSCd )
			if @row> 0 
			begin				select @errmsg= 'Cannot Change Bag Type as Used In Order Component Table'
				goto ErrHndlr
			end
/* *** Jay 2.13(CT) *** */
			
			/****** 'IG BCHR- PValue3 = Blank for primary bags & Components COdes**********/
			select @row= count('x') from inserted where PValue3<> '' and PValue <> 'F'			
			if @row>0 
			begin
				select @errmsg= 'PrdTyp Should Be Enter For Only Flute Bags'
				goto ErrHndlr
			end
			--'IG BCHR-PrdTyp for Flute Bag with RmCtg include D or C		
			select @row= count('x') from inserted where PValue3<> '' and PValue ='F'			
							and not ((','+ PValue1+ ',' like '%,C,%'  ) or (','+ PValue1+ ',' like '%,D,%'  ))
			if @row>0 
			begin
				select @errmsg= 'PrdTyp Should Be Enter For Only Rm category D Or C'
				goto ErrHndlr
			end
			/*** 'IG BCHR-Validation for PValue3		**********/
			/****** PValue3 = blank or Valid PrdTyp for Flute Bags ******/		
			select @row= count('x') from inserted ins where PValue3<> '' and PValue= 'F' and Not(PValue3='WS' or PValue3='GS')																
			if @row> 0 		
			begin		
				select @errmsg= 'PrdTyp Is Invalid'		
				goto ErrHndlr		
			end

				
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End 
		
	/* **** Zubin 211 **** */ 

	/*********Geeta***CRM**2.12******/
		/**** if @PTyp = 'GRCD' - Manali 3.10.0 - Wrong GRCD, GRTYP changed to GRDCD, GRDTYP  */
		if @PTyp = 'GRDCD'
		Begin 
			/****** PMCd = Valid PMCD('GRTYP') ******/
			select @row= count('x') from inserted ins where 
			Not Exists (Select 'x' From Param Pm where Pm.PTyp= 'GRDTYP' and Pm.PMCd=ins.PMCd)
			/*Not Exists (Select 'x' From Param Pm where Pm.PTyp= 'GRTYP' and Pm.PMCd=ins.PMCd)*/
			if @row>0 
			begin
				select @errmsg= 'Grade Type Not Defined'
				goto ErrHndlr
			end

			/****** PSCd <> '' ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'Grade Code Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank******/
			/***** Manali 3.10.0 - PValue as Print Desc - Cannot Be Blank *****/
			select @row=count('x') from inserted ins where PValue = '' /*PValue<>'' */
			if @row>0 
			begin
			  select @errmsg= 'PValue Cannot Be Blank'
				/*select @errmsg= 'PValue Should Be Blank'*/
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
	if @PTyp = 'RMID'
		Begin 
			/****** PMCd = Valid PMCD('GRDRMCTG') ******/
			select @row= count('x') from inserted ins where 
			Not Exists (Select 'x' From Param Pm where Pm.PTyp= 'GRDRMCTG' and Pm.PSCd=ins.PMCd)
			if @row>0 
			begin
				select @errmsg= 'GrRmCtg Not Defined'
				goto ErrHndlr
			end

			/****** PSCd <> '' ******/
			select @row= count('x') from inserted where PSCd= ''
			if @row>0 
			begin
				select @errmsg= 'RmId Cannot Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row=count('x') from inserted ins where PValue<>'' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted where PDesc225<> ''
			if @row>0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
	
	/*********Geeta*************/
		--pg.11
		if @PTyp= 'GSTCD'
		Begin 
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Valid GstBase ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'GSTBASE' 
				and Prm.PMCd= ins.PValue)
			if @row>0 
			begin
				select @errmsg= 'GstBase Not Defined'
				goto ErrHndlr
			end						
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted ins where PDesc225<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		if @PTyp= 'GSTFLDS'
		Begin 
			/****** PMCd = Valid GstCode ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'GSTCD' 
				and Prm.PMCd= ins.PMCd )
			if @row>0 
			begin
				select @errmsg= 'GstCode Not Defined'
				goto ErrHndlr
			end
			/****** PSCd = Valid GstSr ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'GSTSR' 
				and Prm.PMCd= ins.PSCd)
			if @row>0 
			begin
				select @errmsg= 'GstSr Not Defined'
				goto ErrHndlr
			end
			/****** PDesc <> Blank ******/
			select @row= count('x') from inserted ins where PDesc= '' 
			if @row> 0 
			begin
				select @errmsg= 'Description Should Not Be Blank'
				goto ErrHndlr
			end
			/******6.1B PValue = Valid GstOn ******/	
			--select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'GSTON' 
			--	and Prm.PMCd= ins.PValue)
			--if @row>0 
			--begin
			--	select @errmsg= 'GstOn Not Defined'
			--	goto ErrHndlr
			--end									
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted ins where PDesc225<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue1 = Valid GstFor ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'GSTFOR' 
				and Prm.PMCd= ins.PValue1 )
			if @row>0 
			begin
				select @errmsg= 'GstFor Not Defined'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		--vk.19
		if @PTyp= 'SUBCUST'
		Begin 
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PDesc225 = Blank ******/
			select @row= count('x') from inserted ins where PDesc225<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PDesc225 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue = Blank ******/
			select @row= count('x') from inserted ins where PValue<>'' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end						
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end
			
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End
		--6.1B-26			
		if @PTyp= 'USRCMCD'
		Begin 
			/****** PMCd = Valid User ******/
			select @row= count('x') from inserted ins where not exists (select 'x' from Param Prm where Prm.PTyp= 'USR' 
				and Prm.PMCd= ins.PMCd )
			if @row>0 
			begin
				select @errmsg= 'User Code Not Defined'
				goto ErrHndlr
			end
			/****** PSCd = Blank ******/
			select @row= count('x') from inserted where PSCd<> ''
			if @row>0 
			begin
				select @errmsg= 'Sub Code Should Be Blank'
				goto ErrHndlr
			end
			/****** PDesc = Blank ******/
			select @row= count('x') from inserted ins where PDesc<> '' 
			if @row> 0 
			begin
				select @errmsg= 'PDesc Should Be Blank'
				goto ErrHndlr
			end			
			/****** PValue = Blank ******/
			select @row= count('x') from inserted ins where PValue<>'' 
			if @row>0 
			begin
				select @errmsg= 'PValue Should Be Blank'
				goto ErrHndlr
			end						
			/****** PNum = 0 ******/
			select @row= count('x') from inserted where PNum<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum Should Be Zero'
				goto ErrHndlr
			end			
			/****** PValue1 = Blank ******/
			select @row= count('x') from inserted where PValue1<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue1 Should Be Blank'
				goto ErrHndlr
			end
			/****** PNum1 = 0 ******/
			select @row= count('x') from inserted where PNum1<> 0
			if @row>0 
			begin
				select @errmsg= 'PNum1 Should Be Zero'
				goto ErrHndlr
			end
			/****** PValue2 = Blank ******/
			select @row= count('x') from inserted where PValue2<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue2 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValue3 = Blank ******/
			select @row= count('x') from inserted where PValue3<> ''
			if @row>0 
			begin
				select @errmsg= 'PValue3 Should Be Blank'
				goto ErrHndlr
			end
			/****** PValidYN = Blank ******/
			select @row= count('x') from inserted where PValidYN<> ''
			if @row>0
			begin
				select @errmsg= 'PValidYN Should Be Blank'
				goto ErrHndlr
			end
		End

	END

	if @mode in ('D')
	BEGIN
		/*** In Use Checks ***/
		/*** In Use Checks ***/
		/*** In Use Checks ***/
		/*** In Use Checks ***/
		/*** In Use Checks ***/
		/*** In Use Checks ***/

		if @PTyp= 'BCHR' 
		Begin 
			/****** Cannot delete Bag Chr in use in Head table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from Head where hBChr= PMCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Bag Chr In Use In Head Table'
				goto ErrHndlr
			end
			/****** Cannot delete Bag Chr in use in Bag table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from Bag where BChr= PMCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Bag Chr In Use In Bag Table'
				goto ErrHndlr
			end
		/* *** Jay 2.13(CT) *** */
			/****** Cannot delete Bag Chr in use in DsgMst table ******/
			select @row= count('x') from deleted del where PValue = 'C' and 
				exists (select 'x' from DsgMst where DmTcTyp = 'PM' and DmSz= PMCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Code Is Used As Component Code in Design Master Table'
				goto ErrHndlr
			end
			/****** Cannot delete Bag Chr in use in DsgCt table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from DsgCT where DctTcTyp= 'DM' and DctDmSz= '' and DctChr= PMCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Bag Chr In Use In Dsg Component Table'
				goto ErrHndlr
			end
			/****** Cannot delete Bag Chr in use in OrdCt table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from OrdCT where OctCtChr= del.PMCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Bag Chr In Use In Order Component Table'
				goto ErrHndlr
			end
			select @row= count('x') from deleted del where 
				exists (select 'x' from vParam where vPTyp='CHKTOL' and charindex(','+del.PMCd+',',','+vPDesc+',',0)>0)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Bag Chr In Use In CHKTOL definition'
				goto ErrHndlr
			end				

		/* *** Jay 2.13(CT) *** */
		End


		if @PTyp= 'CPYRT' 
		Begin 
			/**** Bef 3.03 - Following check no more reqd for PmCd= 'GPSDCXAM' ****/
			/****** Cannot delete or change the Record of PmCd= 'GPSDCXAM' ******/
			/*
			select @row= count('x') from deleted del where PMCd= 'GPSDCXAM'
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete The Record ''GPSDCXAM'' '
				goto ErrHndlr
			end
			*/
			
			/* **** Manali 3.03 - 12/07/08  */
			/****** Cannot delete or change the Record of PmCd= 'GPSLDCXM' ******/
			select @row= count('x') from deleted del where PMCd= 'GPSLDCXM'
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete The Record ''GPSLDCXM'' '
				goto ErrHndlr
			end
			/* **** Manali 3.03 - 12/07/08  */

			/****** Cannot delete Order Rm Copy Rate Option in use in Head table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where HOrdCpyRtOpt= PMCd or HOdGmChk= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Order Rm Copy Rate Option Is Used In Head Table'
				goto ErrHndlr
			end

			/****** Cannot delete Rate Option in use in OrdDsg table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdGmChk= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Option Is Used In OrdDsg As Gram Check Option'
				goto ErrHndlr
			end

		End

		if @PTyp= 'CURNCY' 
		Begin 
			/* Zubin 213 Cannot delete Currency Code if in use in Head table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from Head where HCd= 'ZSELF' and HBaseCurCd= PMCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Currency Code Is Used In Head Table'
				goto ErrHndlr
			end
			/* Zubin 213 Cannot delete Currency Code if in use in CustMst table (CmCtg= 'T' removed) ******/ 
			select @row= count('x') from deleted del where 
				exists (select 'x' from CustMst where CmCurCd= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Currency Code Is Used In CustMst Table'
				goto ErrHndlr
			end
			/* Zubin 213 Cannot delete Currency Code if in use in Date table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from DtTbl where DtCurCd1= PMCd or DtCurCd2= PMCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Currency Code Is Used In Date Table'
				goto ErrHndlr
			end
			/* Zubin 213 Cannot delete Currency Code if in use in RmRt table (RrTcTyp= 'RM' removed) ******/ 
			select @row= count('x') from deleted del where 
				Exists (select 'x' from RmRt where RrCmCurCd= PMCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Currency Code Is Used In RM Rate Mater'
				goto ErrHndlr
			end
			/* Zubin 213 Cannot delete Currency Code if in use in RmRtHist table ******/
			select @row= count('x') from deleted del where 
				Exists (select 'x' from RmRtHist where RhCmCurCd= PMCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Currency Code Is Used In RM Rate History Mater'
				goto ErrHndlr
			end
			/* Cannot delete Currency Code if in use in LabRt table ******/
			select @row= count('x') from deleted del where 
				Exists (select 'x' from LabRt where LrCmCurCd= PMCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Currency Code Is Used In Labour Rate Mater'
				goto ErrHndlr
			end
			/* Zubin 213 Cannot delete Currency Code if in use in OrdMst (PDC) table ******/
			select @row= count('x') from deleted del where 
				Exists (select 'x' from OrdMst where 
					OmTc In (Select TC.PmCd from Param TC where TC.PTyp= 'TC' and TC.PSCd= '' and 
						TC.PValue= 'PL') and OmCmCurCd= del.PMCd) 

			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Currency Code Is Used In Party Design Combination Master' 
				goto ErrHndlr
			end
			/* Zubin 213 Cannot delete 'RS' Currency Code ******/
			select @row= count('x') from deleted del where PMCd= 'RS' 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete ''RS'' Currency'
				goto ErrHndlr
			end
			/**** 6.1.11 Cannot delete Currency Code if in use in Shipment Note table ******/
			select @row= count('x') from deleted del where 
				Exists (select 'x' from ShpNote where SnCurCd = PMCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Currency Code Is Used In ShpNote Table'
				goto ErrHndlr
			end
		End

		-- 4.1.2.0
		if @PTyp= 'DMMATX' 
		Begin 
			/****** Cannot delete Design Analysis Cd Option in use in DsgAna Table ******/
			select @row= count('x') from deleted del Where PTyp='DMMATX'
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete Design Matrix Entries'
				goto ErrHndlr
			end
		End

		if @PTyp= 'DAANACD' 
		Begin 
			/****** Cannot delete Design Analysis Cd Option in use in DsgAna Table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from DsgAna where DaAnaSr= PMCd and DaAnaCd= PSCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Dsg Analysis Cd In Use In DsgAna Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'DAANAFLD' 
		Begin 
			/****** Cannot delete Design Analysis Field Option in use in Param('DAANACD') ******/ 
			select @row= count('x') from deleted del where 
				exists (select 'x' from Param Pm where Pm.PTyp= 'DAANACD' and Pm.PmCd= del.PMCd) 
			if @row> 0 
			begin 
				select @errmsg= 'Cannot Delete Analysis Fld Definition As Analysis Cd Exists' 
				goto ErrHndlr 
			end 

			/****** Cannot delete Design Analysis Desc Field Option (PValue= 'N') in use in DsgAna ******/ 
			select @row= count('x') from deleted where PValue= 'N' and 
				exists (select 'x' from DsgAna where DaAnaSr= PMCd) 
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete Analysis Fld As Analysis Desc Exists In DsgAna Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'DEFCLR' 
		Begin 
			/****** Cannot delete DEFCLR record ******/ 
			select @row= count('x') from deleted del where PTyp= 'DEFCLR' 
			if @row> 0 
			begin 
				select @errmsg= 'Cannot Delete Default Color Record As Used In System' 
				goto ErrHndlr 
			end 
		End

		if @PTyp= 'DMCOL' 
		Begin 
			/****** Cannot delete Dsg Colour  in use in CustAly table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from CustAly where CaDmCol= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Colour In Use In CustAly Table'
				goto ErrHndlr
			end
			/****** Cannot delete Dsg Colour  in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmCol= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Colour In Use In DsgMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Dsg Colour  in use in OrdDsg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdDmCol= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Colour In Use In OrdDsg Table'
				goto ErrHndlr
			end		End

		if @PTyp= 'DMCTG' 
		Begin 
			/****** Cannot delete Dsg Category in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmCtg= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Category In Use In DsgMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Dsg Category in use in RmRt table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from RmRt where RrDmCtg= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Category In Use In RmRt Table'
				goto ErrHndlr
			end
		End

		/* Sachin 4.1.0.0 */
		if @PTyp = 'CMGRP' 
		Begin 
			/****** Cannot delete Dsg Category in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from CustMst where CmGrp= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Customer Group In Use In CustMst Table'
				goto ErrHndlr
			end
			/* 578-12 Cannot delete group in use in DmAllow table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DmAllow where DmwCmGrp = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Customer Group In Use In DmAllow Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'CMTYP' 
		Begin 
			/****** Cannot delete Dsg Category in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from CustMst where CmTyp= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Customer Type In Use In CustMst Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'PORTCD' 
		Begin 
			select @row=count('x') from deleted del where 
				exists (select 'x' from CustMst where CmPOD= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Port Code In Use In CustMst Table'
				goto ErrHndlr
			end

			select @row=count('x') from deleted del where 
				exists (select 'x' from InvHd where InCmPOD= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Port Code In Use In InvHd Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'CONTRYCD' 
		Begin 
			/****** Cannot delete Dsg Category in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from CustMst where CmFinDstCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Destination Code In Use In CustMst Table'
				goto ErrHndlr
			end

			select @row=count('x') from deleted del where 
				exists (select 'x' from InvHd where InCmFinDstCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Destination Code In Use In InvHd Table'
				goto ErrHndlr
			end
		End
		/* 4.1.0.0 - Chgs End */	

		if @PTyp= 'DMSZ' 
		Begin 
		/****** Cannot delete Dsg Size in use in DsgMst table ******/
			/* *** Jay 3.2.0 [or DmDefSz= PMCd] *** */ 
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmSz= PMCd or DmDefSz= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Size In Use In DsgMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Dsg Size in use in OrdDsg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdDmSz= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Size In Use In OrdDsg Table'
				goto ErrHndlr
			end
		/****** Cannot delete Dsg Size in use in Bag table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from Bag where BOdDmSz= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Size In Use In Bag Table'
				goto ErrHndlr
			end
			/****** Cannot delete Dsg Size in use in Fgd (TcTyp= 'FR') table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from Fgd where FdTc in (Select PMCd from Param where PTyp= 'TC' and PValue= 'FR') 
					and FdDmSz= del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Size In Use In Fgd (Fin Goods Rm) Table'
				goto ErrHndlr
			end
		
		/****** Cannot delete Dsg Size in use in MultiPrcDsg table ******/		/***** Manali 3.10.0 - Multi Price Quotation ****/			select @row=count('x') from deleted del where 
				exists (select 'x' from MultiPrcQtDsg where MqdDmSz= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Size In Use In MultiPrcQtDsg Table'
				goto ErrHndlr
			end
			/****** pg.5 Cannot delete Dsg Size in use in RmRt table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from RmRt where RrDmSz = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Size In Use In RmRt Table'
				goto ErrHndlr
			end

		End

		if @PTyp= 'GLDAS' 
		Begin 
			/****** Cannot delete Gold As Option in use in Head table ******/
			/* Sachin 2.14.0 - HNegStkRmCtg added below */
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where HCmGldAs= PMCd Or HNegStkRmCtg=PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Gold As Option Is Used In Head Table'
				goto ErrHndlr
			end

			/****** Cannot delete Gold As Option in use in CustMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from CustMst where CmGldAs= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Gold As Option Is Used In CustMst Table'
				goto ErrHndlr
			end

			/****** Cannot delete Gold As Option in use in OrdMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdMst where OmGldAs= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Gold As Option Is Used In OrdMst Table'
				goto ErrHndlr
			end

			/****** Cannot delete Gold As Option in use in OrdDsg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdGldAs= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Gold As Option Is Used In OrdDsg Table'
				goto ErrHndlr
			end

			/****** Cannot delete Gold As Option in use in InvDsg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from InvDsg where IdGldAs= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Gold As Option Is Used In InvDsg Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'IMPTYP' 
		Begin 
			/****** Cannot delete Import Type in use in Txn Of Type 'TZ' ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from Txn where TImpTyp= PMCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Import Type In Use In Txn Table'
				goto ErrHndlr
			end
		End

	/*	if @PTyp= 'LOCKDT' 
		Begin 
			/****** Cannot delete LOCKDT In a Purged Database (Sachin) ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from Param where PTyp= 'PURGEDT' And PMCd='PURGEDT' And Cast(PDesc225 as smalldatetime)>'01/01/80') 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete LOCKDT In a Purged Database'
				goto ErrHndlr
			end
		End
*/
		if @PTyp= 'IWTFRORD' 
		Begin 
			/****** Cannot delete Weight From Ord Option in use in Head table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where HCmIWtFrOrd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Weight From Order Option Is Used In Head Table'
				goto ErrHndlr
			end
			/****** Cannot delete Weight From Ord Option in use in CustMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from CustMst where CmIWtFrOrd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Weight From Order Option Is Used In CustMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Weight From Ord Option in use in OrdMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdMst where OmIWtFrOrd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Weight From Order Option Is Used In OrdMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Weight From Ord Option in use in OrdDsg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdIWtFrOrd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Weight From Order Option Is Used In OrdDsg Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'KT' 
		Begin 
			/****** Cannot delete Karat in use in RmMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from RmMst where RmKt= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Karat In Use In RmMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Karat in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmKt= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Karat In Use In DsgMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Karat in use in OrdDsg table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdKt= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Karat In Use In OrdDsg Table'
				goto ErrHndlr
			end
			/****** Cannot delete Karat in use in InvDsg table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from InvDsg where IdKt= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Karat In Use In InvDsg Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'LABMCD' 
		Begin 
			/****** Cannot delete 'SET' main code ******/
			select @row=count('x') from deleted del where PMCd= 'SET' 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete ''SET'' Lab Main Cd As It Is Integral To The Working Of The System'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Main Code in use in Param table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Param Prm where Prm.PTyp= 'LABSCD' and Prm.PMCd= del.PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Lab Main Code Is Used In Param (LABSCD) Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Main Code in use in DsgLab table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgLab where DlMCd= PMCd)
			if @row> 0 
			Begin
				select @errmsg='Cannot Delete As Lab Main Code Is Used In DsgLab Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Main Code in use in LabRt table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from LabRt where LrMCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Lab Main Code Is Used In LabRt Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Main Code in use in OrdLab table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdLab where OlMCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Lab Main Code Is Used In OrdLab Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Main Code in use in InvLab table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from InvLab where iLMCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Lab Main Code Is Used In InvLab Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Main Code in use in RmRt table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from RmRt where RrLabMCd= PMCd)
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Lab Main Code Is Used In RmRt Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Main Code in use in RmRtHist table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from RmRtHist where RhLabMCd= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Lab Main Code Is Used In RmRtHist Table'
				goto ErrHndlr
			end
			/****** Manali 3.10.0 - Cannot delete Lab Main Code in use in MultiPrcQtLab table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from MultiPrcQtLab where MqlMCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Lab Main Code Is Used In MultiPrcQtLab Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'LABSCD' 
		Begin 
			/****** Cannot delete Lab Sub Code in use in DsgLab table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgLab where DlMCd= PMCd and DlSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Labour Sub Code In Use In DsgLab Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Sub Code in use in LabRt table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from LabRt where LrMCd= PMCd and LrSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Labour Sub Code In Use In LabRt Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Sub Code in use in OrdLab table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdLab where OlMCd= PMCd and OlSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Labour Sub Code In Use In OrdLab Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Sub Code in use in DsgRm table ******/
			select @row=count('x') from deleted del where PMCd= 'SET' and 
				exists (select 'x' from DsgRm where DrSetSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Setting Code In Use In DsgRm Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Sub Code in use in OrdRm table ******/
			select @row=count('x') from deleted del where PMCd= 'SET' and 
				exists (select 'x' from OrdRm where OrSetSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Setting Code In Use In OrdRm Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Sub Code in use in InvRm table ******/
			select @row=count('x') from deleted del where PMCd= 'SET' and 
				exists (select 'x' from InvRm where IrSetSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Setting Code In Use In InvRm Table'
				goto ErrHndlr
			end
			/****** Cannot delete Lab Sub Code in use in InvLab table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from InvLab where iLMCd= PMCd and iLSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Labour Sub Code In Use In InvLab Table'
				goto ErrHndlr
			end
			/* *** Jay 2.14Next *** */
			/****** Cannot delete Lab Sub Code in use in PrdPts table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from PrdPts where PpSetTyp= PSCd)
			if @row>0 
				Begin
					select @errmsg='Cannot Delete As Labour Sub Code In Use In PrdPts Table'
					goto ErrHndlr
				end
			/* *** Jay 2.14Next *** */
			/****** Manali 3.10.0 - Cannot delete Lab Sub Code in use in MultiPrcQtRm table ******/
			select @row=count('x') from deleted del where PMCd= 'SET' and 
				exists (select 'x' from MultiPrcQtRm where MqrSetSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Setting Code In Use In MultiPrcQtRm Table'
				goto ErrHndlr
			end
			/****** Manali 3.10.0 - Cannot delete Lab Sub Code in use in MultiPrcQtLab table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from MultiPrcQtLab where MqLMCd= PMCd and MqLSCd= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Labour Sub Code In Use In MultiPrcQtLab Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'LSCTG' 
		Begin 
			/****** Cannot delete Loss Ctg in use in DsgMst table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmLsCtg= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Loss Ctg In Use In DsgMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Loss Ctg in use in RmRt table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmRt where RrTcTyp= 'LS' and RrLsCtg= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Loss Ctg In Use In RmRt Table For ''LS'' TcTyp'
				goto ErrHndlr
			end
		End

		if @PTyp= 'PARENTCD' 
		Begin 
			/****** Cannot delete Parent Cd in use in RmMst table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmMst where RmCtg= PValue and RmSCtg= PValue1 and RmParentCd= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Parent Cd Is Used In RmMst Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'PCT' 
		Begin 
			/****** Cannot Delete Records Of Ptyp= 'PCT' *****/
			select @errmsg= 'Cannot Delete Records Of Param Type ''PCT'' '
			goto ErrHndlr
		End

    /***** Manali 3.9.0 - Output Folde For Report *****/ 
		if @PTyp= 'OUTPUT' 
		Begin 
			/****** Cannot Delete Records Of Ptyp= 'Output' *****/
			select @errmsg= 'Cannot Delete Records Of Param Type ''OUTPUT'' '
			goto ErrHndlr
		End
    /***** Manali 3.9.0 - Output Folde For Report *****/ 

		if @PTyp= 'PLCHR' 
		Begin 
			/****** Cannot delete Packing List Chr  in use in Fgd table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Fgd where FdPlChr= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Packing List Chr In Use In Fgd Table'
				goto ErrHndlr
			end
			/****** Cannot delete Packing List Chr  in use in Head table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where hFgPlChr= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Packing List Chr In Use In Head Table'
				goto ErrHndlr
			end		End

		if @PTyp= 'PRDCTG' 
		Begin 
			/****** Cannot delete Production Ctg in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmPrdCtg= PMCd)
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Production Category In Use In DsgMst Table'
				goto ErrHndlr
			end		End

		/**********Bhavna************/
		Begin 
			/****** Cannot delete Production Ctg in use in DsgPrd table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgPrd where DppPrdCtg= PMCd)
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Production Category In Use In DsgMst Table'
				goto ErrHndlr
			end		End
		
		/*** (Jen 2.14 Next) ***/
		if @PTyp= 'PRDSTG' 
		Begin 
			/****** Cannot delete Production Stage ******/
			select @row= count('x') from deleted del 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete Production Stage'
				goto ErrHndlr
			end
		End 
		/*** (Jen 2.14 Next) ***/

		if @PTyp= 'PRTCD' 
		Begin 
			/****** Cannot delete Priority Code in use in Head table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where hOmPrtCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Priority Code In Use In Head Table'
				goto ErrHndlr
			end
			/****** Cannot delete Priority Code in use in OrdDsg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdPrtCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Priority Code In Use In OrdDsg Table'
				goto ErrHndlr
			end
			/****** Cannot delete Priority Code in use in OrdMst table ******/			select @row=count('x') from deleted del where 
			exists (select 'x' from OrdMst where OmPrtCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Priority Code In Use In OrdMst Table'
				goto ErrHndlr
			end		End

		if @PTyp= 'REGION' 
		Begin 
			/****** Cannot delete Region Code in use in CustMst table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from CustMst where CmRegnCd= PMCd)
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete As Region Code In Use In CustMst Table'
				goto ErrHndlr
			end
			/* 578-12 Cannot delete region in use in DmAllow table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DmAllow where DmwReg = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Region Code In Use In DmAllow Table'
				goto ErrHndlr
			end
			
		End

		if @PTyp= 'REJCD' 
		Begin 
			/****** Cannot delete Rejection Code in use in Txnd table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from Txnd where TdRjCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Customer Category In Use In Txnd Table'
				goto ErrHndlr
			end		End

		/*** (Jen 3.01.1) ***/
		if @PTyp= 'REMAKE' 
		Begin 
			/****** Cannot delete Remake Option in use in InvHd table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from InvHd where InRemakingCtgs= PMCd)
			if @row> 0 
			Begin
				select @errmsg='Cannot Delete As Remake Option Is Used In InvHd table'
				goto ErrHndlr
			end
		End
		/*** (Jen 3.01.1) ***/
		/**** Jay 3.1.0_1 ****/
		if @PTyp= 'SETCD' 
		Begin 
			/****** Cannot delete Remake Option in use in InvHd table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmSetCd= PMCd)
			if @row> 0 
			Begin
				select @errmsg='Cannot Delete As Set Code Is Used In Design Master'
				goto ErrHndlr
			end

			/****** PDesc225 <> Blank ******/
			select @row= count('x') from Inserted where PDesc225= ''
			if @row> 0 
			Begin
				select @errmsg= 'Description Cannot Be Blank'
				goto ErrHndlr
			End
		End
		/**** Jay 3.1.0_1 ****/
		/*** (Jen 2.13) ***/
		if @PTyp= 'RETCHR' 
		Begin 
			/****** Cannot delete RETCHR in use in RetMem table ******/			select @row=count('x') from deleted del where 
				exists (select 'x' from RetMem where ReChr= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As ''Return Memo Character'' In Use In RetMem Table'
				goto ErrHndlr
			end		End
		/*** (Jen 2.13) ***/

		if @PTyp= 'RMCOL' 
		Begin 
			/****** Cannot delete Raw Material Colour in use in RmMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from RmMst where RmCtg= PMCd and RmCol= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Rm Colour In Use In RmMst Table'
				goto ErrHndlr
			end	
			/*VK.5*/
			/****** Cannot delete Raw Material Colour in use in CertDet table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from CertDet where CdRmCtg= PMCd and CdRmCol= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Rm Colour In Use In CertDet Table'
				goto ErrHndlr
			end		
		End

		if @PTyp= 'RMSCTG' 
		Begin 
			/****** Cannot delete Rm Sub Ctg in use in Param(PARENTCD) table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from Param Pm where Pm.PTyp= 'PARENTCD' and Pm.PValue= del.PMCd and Pm.PValue1= del.PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In Param(''PARENTCD'') Table'
				goto ErrHndlr
			end
			/****** Cannot delete Rm Sub Ctg in use in RmAlyCmp table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmAlyCmp where RaCmpCtg= PMCd and RaCmpSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In RmAlyCmp Table'
				goto ErrHndlr
			end
			/****** Cannot delete Rm Sub Ctg in use in RmMst table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmMst where RmCtg= PMCd and RmSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In RmMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Rm Sub Ctg in use in DsgRm table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from DsgRm where DrRmCtg= PMCd and DrRmSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In DsgRm Table'
				goto ErrHndlr
			end
			/****** Cannot delete Rm Sub Ctg in use in OrdTgt table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from OrdTgt where OtRmCtg= PMCd and OtRmSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In OrdTgt Table'
				goto ErrHndlr
			end
			/****** Cannot delete Rm Sub Ctg in use in OrdRm table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from OrdRm where OrRmCtg= PMCd and OrRmSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In OrdRm Table'
				goto ErrHndlr
			end
			/****** Cannot delete Rm Sub Ctg in use in RmRt table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmRt where RrCtg= PMCd and RrSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In RmRt Table'
				goto ErrHndlr
			end
			/****** Cannot delete Rm Sub Ctg in use in RmRtHist table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmRtHist where RhCtg= PMCd and RhSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In RmRtHist Table'
				goto ErrHndlr
			end
			/****** Cannot delete Rm Sub Ctg in use in TxndPrd table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from TxndPrd where TpRmCtg= PMCd and TpRmSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In TxndPrd Table'
				goto ErrHndlr
			end
			/****** Manali 3.10.0 - Cannot delete Rm Sub Ctg in use in MultiPrcQtRm table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from MultiPrcQtRm where MqrRmCtg= PMCd and MqrRmSCtg= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Rm Sub Ctg Is Used In MultiPrcQtRm Table'
				goto ErrHndlr
			end
			/****** mw.98 Cannot delete Rm Sub Ctg in use in vParam table ******/
			select @row=count('x') from deleted del where
				exists (select 'x' from vParam where vPTyp= 'CHKTOL' and charindex(','+PSCd+',', ','+vPValue2+',')> 0)
			if @row>0
			Begin
				select @errmsg='Cannot Delete As Rm Sub Ctg In Use In vParam Table Where vPTyp= ''CHKTOL'' '
				goto ErrHndlr
			end

		End


		/*** (Jen 2.14 Next) ***/
		if @PTyp= 'RTC' 
		Begin 
			/****** Cannot delete RTC record ******/ 
			select @row= count('x') from deleted del where PTyp= 'RTC' and PMCd='RTC' and PSCd=1
			if @row> 0 
			begin 
				select @errmsg= 'Cannot Delete Report Time Control Record where SrNo is 1' 
				goto ErrHndlr 
			end 
		End
		/*** (Jen 2.14 Next) ***/

		/* ****** Sachin - Multi Server ****** */
		if @PTyp= 'REPCTR' 
		Begin 
			/****** Cannot delete RTC record ******/ 
			select @row= count('x') from deleted del where PTyp= 'REPCTR' 
			if @row> 0 
			begin 
				select @errmsg= 'Cannot Delete Replication Counter Record' 
				goto ErrHndlr 
			end 
		End
		/* ****** Sachin - Multi Server ****** */

		if @PTyp= 'SALCTG' 
		Begin 
			/*******Geeta*****Crm***2.12*****Checked DmCtg=PMCd and DmSalCtg=PSCd****/
			/****** Cannot delete Design Sales Category in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmCtg=PMCd and DmSalCtg= PSCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Design Sales Category In Use In DsgMst Table'
				goto ErrHndlr
			end		End


		/*** (Jen 2.14 Next) ***/
		if @PTyp= 'SEOPWD' 
		Begin 
			/****** Cannot Delete Record Of Type 'SEOPWD' ******/
			select @errmsg= 'SEO Passowrd Record Cannot Be Deleted' 
			goto ErrHndlr
		End
		/*** (Jen 2.14 Next) ***/ 


		if @PTyp= 'SP' 
		Begin 
			/****** Cannot delete Sales Person Code in use in CustMst ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from CustMst where CmSalPer= PMCd) 
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Sales Person Code In Use In CustMst Table'
				goto ErrHndlr
			end
			/*pg3a*/
			/****** Cannot delete Sales Person in use in OrdMst ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdMst where OmCmSalPer= PMCd )
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Sales Person Code In Use In OrdMst Table'
				goto ErrHndlr
			end		
			
		End

		if @PTyp= 'STWGRP' 
		Begin 
			/****** Cannot delete Size To Weight Group in use in RmMst table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmMst where RmSTWGrp= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Size To Weight Group In Use In RmMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Size To Weight Group in use in RmRt table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmRt where RrTcTyp= 'STW' and RrSTWGrp= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Size To Weight Group In Use In RmRt Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'SUBPRC' 
		Begin 
/*  #################################  Manoj #### Ver: 2.0.6  #################################  */
			/****** Cannot delete Sub Process in use in TxndPrd table ******/
			select @row=count('x') from deleted Del where 
				exists (select 'x' from PrdPts where PpLoc = Del.PMCd And PpSubPrc = Del.PSCd)
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Sub Process In Use In PrdPts Table'
				goto ErrHndlr
			end			/****** Cannot delete Sub Process in use in TxndPrd table ******/
			select @row=count('x') from deleted del where 
				exists (select TpSubPrc from TxndPrd Join  Txnd ON TdTc= TpTc and TdYy= TpYy and TdChr= TpChr and TdNo= TpNo 
						and TdSr= TpSr and TdSrNo= 0 And TdByLoc = Del.PMCd And  TpSubPrc= Del.PSCd )
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Sub Process In Use In TxndPrd Table'
				goto ErrHndlr
			end
/*			select @row=count('x') from deleted del where 
				exists (select 'x' from TxndPrd where IsNull((Select TdByLoc from Txnd where TdTc= TpTc and TdYy= TpYy 
					and TdChr= TpChr and TdNo= TpNo and TdSr= TpSr and TdSrNo= 0),'') = PMCd  and TpSubPrc= Del.PSCd)
*/		End
/*  #########################################  Manoj  #########################################  */

/****************Bhavna RWRK DELETE Valid***********/
		if @PTyp= 'RWRK' 
		Begin 

			/****** Cannot delete RWRK in use in TxndRwrk table ******/
			
			select @row=count('x') from deleted del where 
				exists (select TrwRwrkCd from TxndRwrk Join  Txnd ON TdTc= TrwTc and TdYy= TrwYy and TdChr= TrwChr and TdNo= TrwNo 
						and TdSr= TrwSr and TdSrNo= 0 And TdByLoc = Del.PMCd And  TrwRwrkCd = Del.PSCd )
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As ReWork Code In Use In TxndRwrk Table'
				goto ErrHndlr
			end
		End
/*******************Bhavna RWRK DELETE Valid**********/			
		/* InUse Check */
		if @PTyp= 'CELL' 
		Begin 
			/****** Cannot delete Cell in use in vParam table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from vParam where vPTyp ='PRDSEQ' and vPValue2 = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Cell In Use In Production Sequence'
				goto ErrHndlr
			end
			/****** Cannot delete Cell in use in OrdLn table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdLn where OdlCell = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Cell In Use In OrdLn Table'
				goto ErrHndlr
			end			
			/****** Cannot delete Cell in use in OrdMst table ******/			
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdMst where OmCell = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Cell In Use In OrdMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Cell in use in Head table ******/			
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where hCell = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Cell In Use In Head Table'
				goto ErrHndlr
			end
			/****** Cannot delete Cell in use in CustMst table ******/			
			select @row=count('x') from deleted del where 
				exists (select 'x' from CustMst where CmCell = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Cell In Use In CustMst Table'
				goto ErrHndlr
			end	
			/****** Cannot delete Cell in use in OrdDsg table ******/			
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdCell = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Cell In Use In OrdDsg Table'
				goto ErrHndlr
			end		
		End
		/* InUse Check */
		/****** Cannot delete DTCYC in use in OrdLn table ******/
		if @PTyp= 'DTCYC' 
		Begin 						
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdLn where OdlDtCyc = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As DtCyc In Use In OrdLn Table'
				goto ErrHndlr
			end						
		End
		/****** Cannot delete Skill in use in Production Point table ******/
		if @PTyp= 'SKL' 
		Begin 						
			select @row=count('x') from deleted del where 
				exists (select 'x' from PrdPts where PpSkl = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Skill In Use In PrdPts Table'
				goto ErrHndlr
			end						
		End
		/****** Cannot delete PrdLine in use in vParam table ******/
		if @PTyp= 'PRDLINE' 
		Begin 			
			select @row=count('x') from deleted del where 
				exists (select 'x' from vParam where vPTyp ='PRCSSEQ' and vPValue1 = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As PrdLine In Use In vParam whee vPtyp = PRCSSEQ '
				goto ErrHndlr
			end
		End
	/****** Cannot delete PYMTRM in use in CustMst table ******/
		if @PTyp= 'PYMTRM' 

		Begin 						








			select @row=count('x') from deleted del where 
				exists (select 'x' from CustMst where CmPayTermsCd = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As PayTermCode In Use In CustMst Table'
				goto ErrHndlr
			end						
		End	
		/* InUse Check */
		/* IG.44 ***/
		if @PTyp= 'BMRSNCD' 
		Begin 			
		/****** Cannot delete the default Record of PmCd= '-' ******/
			select @row= count('x') from deleted del where PMCd= '-'
			if @row> 0 
			begin
				select @errmsg= 'Cannot Delete The Record ''-'' '
				goto ErrHndlr
			end
			/****** Cannot delete B/M Reason Code in use in Txnd table ******/			
			select @row=count('x') from deleted del where 
				exists (select 'x' from Txnd where TdBmRsnCd = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As B/M Reason Code In Use In Txnd Table'
				goto ErrHndlr
			end
		end	
				
		--pg.11 
		if @PTyp= 'GSTBASE' 
		Begin 			
			/****** Cannot delete GstBase in use in Param table ******/		
			select @row=count('x') from deleted del where  
				exists (select 'x' from Param where PTyp='GSTCD' and PValue =del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As GstBase In Use In Param Where PTyp= ''GSTCD'' ' 
				goto ErrHndlr
			end
			/****** Cannot delete GstBase in use in InvHd table ******/		
			select @row=count('x') from deleted del where  
				exists (select 'x' from InvHd where InGstBase =del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As GstBase In Use In InvHd Table ' 
				goto ErrHndlr
			end
		end	
		if @PTyp= 'GSTCD' 
		Begin 			
			/****** Cannot delete GstCode in use in Param table ******/		
			select @row=count('x') from deleted del where  
				exists (select 'x' from Param where PTyp='GSTFLDS' and PMCd =del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As GstCode In Use In Param Where PTyp= ''GSTFLDS'' ' 
				goto ErrHndlr
			end
			/****** Cannot delete GstCode in use in vParam table ******/		
			select @row=count('x') from deleted del where  
				exists (select 'x' from vParam where vPTyp='GSTTCCHR' and vPValue =del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As GstCode In Use In vParam Where vPTyp= ''GSTTCCHR'' ' 
				goto ErrHndlr
			end
			/****** Cannot delete GstCode in use in InvHd table ******/		
			select @row=count('x') from deleted del where  
				exists (select 'x' from InvHd where InGstCd =del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As GstCode In Use In InvHd Table ' 
				goto ErrHndlr
			end
		end	
		if @PTyp= 'GSTSR' 
		Begin 			
			/****** Cannot delete GstSr in use in Param table ******/		
			select @row=count('x') from deleted del where  
				exists (select 'x' from Param where PTyp='GSTFLDS' and PSCd =del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As GstSr In Use In Param Where PTyp= ''GSTFLDS'' ' 
				goto ErrHndlr
			end
		end	
		if @PTyp= 'GSTON' 
		Begin 			
			/******6.1B Cannot delete GstOn in use in Param table ******/		
			--select @row=count('x') from deleted del where  
			--	exists (select 'x' from Param where PTyp='GSTFLDS' and PValue =del.PMCd)
			--if @row>0 
			--Begin
			--	select @errmsg='Cannot Delete As GstOn In Use In Param Where PTyp= ''GSTFLDS'' ' 
			--	goto ErrHndlr	
			--end			
			/****** Cannot delete GstOn in use in InvHd table ******/		
			select @row=count('x') from deleted del where  
				exists (select 'x' from InvHd where InGstOn1=del.PMCd or InGstOn2=del.PMCd or InGstOn3=del.PMCd or InGstOn4=del.PMCd
				or InGstOn5=del.PMCd or InGstOn6=del.PMCd or InGstOn8 =del.PMCd or InGstOn9=del.PMCd or InGstOn10=del.PMCd or InGstOn11=del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As GstOn In Use In InvHd Table ' 
				goto ErrHndlr
			end			
		end	
		if @PTyp= 'GSTFOR' 
		Begin 			
			/****** Cannot delete GstFor in use in Param table ******/		
			select @row=count('x') from deleted del where  
				exists (select 'x' from Param where PTyp='GSTFLDS' and PValue1 =del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As GstFor In Use In Param Where PTyp= ''GSTFLDS'' ' 
				goto ErrHndlr	
			end
		end	

		if @PTyp= 'SUBSHP' 
		Begin 
			
			/****** Cannot delete Sub Shape in use in Param BCHR table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Param where PTyp='BCHR' and PDesc225=del.PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Sub Shape In Use In Param BCHR'
				goto ErrHndlr
			end

			/****** Cannot delete Sub Shape in use in DsgRm table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgRm where DrSubShp= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Sub Shape In Use In DsgRm Table'
				goto ErrHndlr
			end
			/****** Cannot delete Sub Shape in use in OrdRm table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdRm where OrSubShp= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Sub Shape In Use In OrdRm Table'
				goto ErrHndlr
			end				
			/****** mw.98 Cannot delete Sub Shape in use in vParam table ******/
			select @row=count('x') from deleted del where
				exists (select 'x' from vParam where vPTyp= 'CHKTOL' and vPDesc225= PMCd)
			if @row>0
			Begin
				select @errmsg='Cannot Delete As Sub Shape In Use In vParam Table Where vPTyp= ''CHKTOL'' '
				goto ErrHndlr
			end
			/****** 11.b Cannot delete Sub Shape in use in CertDet table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from CertDet where CdSubShp= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Sub Shape In Use In CertDet Table'
				goto ErrHndlr
			end

		End

		/*** (Jen 2.13) (07/11/06) ***/
		/****** Cannot Delete Recs Of This PTyp ******/
		if @PTyp= 'SYSOPT' 
		Begin 
			select @errmsg= 'Cannot Delete Records Of SYSOPT Param Type' 
			goto ErrHndlr
		End
		/*** (Jen 2.13) (07/11/06) ***/

		if @PTyp= 'TC' 
		Begin 
			/****** Cannot Delete Records Of Ptyp= 'TC' *****/
			select @errmsg= 'Cannot Delete Records Of Param Type ''TC'' '
			goto ErrHndlr
		End
		
		/***** Manali 3.8.0 - 'METLR' *****/
		if @PTyp= 'METLR' 
		Begin 
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where HMetlrCd= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Metler Cd In Use In Head Master'
				goto ErrHndlr
			end
		End

		if @PTyp= 'USR' 
		Begin 
			/* Geeta CRM (Bef Emr 213) */ 
			/****** Cannot delete User Code in use in UsrAccess table ******/
/*			select @row=count('x') from deleted del where 
				exists (select 'x' from UsrAccess where UaUsrCd= PMCd)
*/
			/* Geeta CRM Emr 213 */ 
			select @row=count('x') from deleted del where 
				exists (select 'x' from UsrAccess where UaUsrCd= PMCd) and  SessionProperty('replication_agent')<> 1
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As User Code In Use In UsrAccess Table'
				goto ErrHndlr
			end

			Delete From UsrAccess where UaUsrCd in (Select PMCd From deleted where PTyp='USR')  and SessionProperty('replication_agent')= 1 
			/* Geeta CRM */ 

			/* Zubin 212 */
			/****** (Jen 2.14 Next Patch 1) added 'ADM' ******/
			/****** Manali 3.7.1 - 'JMIC' added ******/
			select @row= count('x') from deleted where PMCd= 'MIC' 
			if @row>0
			begin
				select @errmsg= 'Users ''MIC'' Cannot Be Deleted' 
				goto ErrHndlr
			end 
			/* Zubin 212 */

			/* **** Manali 3.8.0 - Metler **** */
			select @row= count('x') from deleted del where 
				exists(Select 'x' from Param P1 Where P1.PTyp='USR' and ','+del.PDesc225+',' like '%,'+P1.PMCd+',%')
			if @row>0
			begin
				select @errmsg= 'Cannot Delete As User Code As It is Assigned Against Metler Code' 
				goto ErrHndlr
			end 

		End


		/*** (Jen 3.01) ***/
		if @PTyp = 'USRGRP'
		Begin 
			/****** Cannot delete User Group in use in Param table (PTyp= 'USR') ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from Param P where P.PTyp= 'USR' and P.PValue3= del.PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As User Group In Use In ''USR'' Parameter'
				goto ErrHndlr
			end
		End
		/*** (Jen 3.01) ***/


		if @PTyp= 'VACTG' 
		Begin 
			/****** Cannot delete Value Addition Category in use in Head table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where HDmVaCtg= PMCd)
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Value Addition Category Is Used In Head Table'
				goto ErrHndlr
			end
			/****** Cannot delete Value Addition Category in use in DsgMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from DsgMst where DmVaCtg= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Value Addition Category Is Used In DsgMst Table'
				goto ErrHndlr
			end
			/****** Cannot delete Value Addition Category in use in OrdDsg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdVaCtg= PMCd)
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Value Addition Category Is Used In OrdDsg Table'
				goto ErrHndlr
			end
			/****** Cannot delete Value Addition Category in use in InvDsg table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from InvDsg where IdVaCtg= PMCd)
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Value Addition Category Is Used In InvDsg Table'
				goto ErrHndlr
			end
		End

/***
		if @PTyp in ('WIWSEQ', 'WIWLONG')  
		Begin 
		End
***/
		/*****Cannot Delete RmId in Use in RmMst**************************/
		if @PTyp= 'RMID' 
		Begin 
			/****** Cannot delete RmId in use in RmMst table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from RmMst where RmCtg= PMCd and  RmId= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As RmId Is Used In RmMst Table'
				goto ErrHndlr
			end
		End

		if @PTyp= 'GRDCD' 
		Begin 
		/****** Cannot delete GRDCD in use in Grade table ******/
			select @row= count('x') from deleted del where 
				exists (select 'x' from Grade where GrTyp= PMCd and  GrCd= PSCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As GrCd Is Used In Grade Table'
				goto ErrHndlr
			end

			select @row= count('x') from deleted del where PMCd='MET' and
				exists (select 'x' from MultiPrcQtMst where PSCd In (MqmMetGrd1, MqmMetGrd2,MqmMetGrd3,MqmMetGrd4)  )
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As GrCd Is Used In MultiPrcQtMst Table'
				goto ErrHndlr
			end

			select @row= count('x') from deleted del where PMCd='DIA' and
				exists (select 'x' from MultiPrcQtMst where 
				PSCd In (MqmDiaGrd1, MqmDiaGrd2,MqmDiaGrd3,MqmDiaGrd4,MqmDiaGrd5,MqmDiaGrd6,MqmDiaGrd7,MqmDiaGrd8))
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As GrCd Is Used In MultiPrcQtMst Table'
				goto ErrHndlr
			end
		End


		if @PTyp= 'YY' 
		Begin 
			/****** Cannot delete Year in use in Bag table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Bag where BYy= PMCd or BLstYy= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Year In Use In Bag Table'
				goto ErrHndlr
			end
			/****** Cannot delete Year in use in Fg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Fg where FgYy= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Year In Use In Fg Table'
				goto ErrHndlr
			end
			/****** Cannot delete Year in use in Head table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Head where hYy= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Year In Use In Head Table'
				goto ErrHndlr
			end
			/****** Cannot delete Year in use in OrdMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdMst where OmYy= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Year In Use In OrdMst Table'
				goto ErrHndlr
			end			/****** Cannot delete Year in use in Txn table ******/ 
			select @row= count('x') from deleted del where 
				exists (select 'x' from Txn where TYy= PMCd)
			if @row> 0 
			Begin
				select @errmsg= 'Cannot Delete As Year In Use In Txn Table'
				goto ErrHndlr
			end
			/****** Cannot delete Year in use in Txnd table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from Txnd where TdRefYy= PMCd or TdPrevYy= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Year In Use In Txnd Table'
				goto ErrHndlr
			end
			/****** Cannot delete Year in use in TxndZ table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from TxndZ where TzYy= PMCd or TzRefYy= PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Year In Use In TxndZ Table'
				goto ErrHndlr
			end
			/****** Cannot delete Year in use in InvHd table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from InvHd where InYy= PMCd)
			if @row>0 
			Begin
				select @errmsg= 'Cannot Delete As Year In Use In InvHd Table'
				goto ErrHndlr
			end
			/****** 6.1.11 Cannot delete Year in use in shipment note table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from ShpNote where SnYy = PMCd )
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Year In Use In ShpNote Table'
				goto ErrHndlr
			end
		End

		-- vk.19 
		if @PTyp= 'SUBCUST' 
		Begin 			
			/****** Cannot delete Sub customer in use in OrdDsg table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from OrdDsg where OdSubCust = PMCd)
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Sub Customer In Use In OrdDsg Table'
				goto ErrHndlr
			end			
		end
		/* JewelSmith.2 */
		if @PTyp= 'STNCLS' 
		Begin 			
			/****** Cannot delete Stone class in use in RmMst table ******/
			select @row=count('x') from deleted del where 
				exists (select 'x' from RmMst where RmCtg =PMCd and RmStnCls = PSCd )
			if @row>0 
			Begin
				select @errmsg='Cannot Delete As Stone Class In Use In RmMst Table'
				goto ErrHndlr
			end			
		end
				
	END


	if @SYS= 'XEMR' 
	BEGIN

/*		if @mode in ('A','C')
		BEGIN
		END
*/


		if @mode in ('D')
		BEGIN
			/*** In Use Checks ***/
			/*** In Use Checks ***/
			/*** In Use Checks ***/
			/*** In Use Checks ***/
			/*** In Use Checks ***/
			/*** In Use Checks ***/

			if @PTyp= 'CURNCY' 
			Begin 
				/****** Cannot delete Currency Cd in use in xParam table where xPTyp= 'BASECUR' ******/
				select @row= count('x') from deleted del where 
					exists (select 'x' from xParam where xPTyp= 'BASECUR' and xPValue= del.PMCd)
				if @row> 0 
				begin
					select @errmsg= 'Cannot Delete As Currency Cd In Use In xParam Table Where xPTyp= ''BASECUR'' ' 
					goto ErrHndlr
				end 
	
				/****** Cannot delete Currency Cd in use in Param table where PTyp= 'BASECUR' */
/*				select @row= count('x') from deleted del where 
					exists (select 'x' from Param Pm where Pm.PTyp= 'BASECUR' 
							and Pm.PValue= del.PMCd)
				if @row> 0 
				begin
					select @errmsg= 'Cannot Delete As Currency Cd In Use In Param Table Where PTyp= 'BASECUR''
					goto ErrHndlr
				end
*/	
				/****** Cannot delete Currency Code in use in xTxn table ******/
				select @row= count('x') from deleted del where 
					exists (select 'x' from xTxn where  xTCurCd= PMCd)
				if @row> 0 
				Begin
					select @errmsg= 'Cannot Delete As Currency Code Is Used In xTxn Table'
					goto ErrHndlr
				end
			
			End

			if @PTyp= 'DMCOL' 
			Begin 
				/****** Cannot delete Dsg Colour  in use in xTxnDsg table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxnDsg where xTdDmCol= PMCd)
				if @row>0 
				Begin
					select @errmsg='Cannot Delete As Design Colour In Use In xTxnDsg Table'
					goto ErrHndlr
				end
			End

			if @PTyp= 'DMSZ' 
			Begin 
				/****** Cannot delete Dsg Size in use in xTxnDsg table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxnDsg where xTdDmSz= PMCd)
				if @row>0 
				Begin
					select @errmsg='Cannot Delete As Design Size In Use In xTxnDsg Table'
					goto ErrHndlr
				end
			End 

			if @PTyp= 'GLDAS' 
			Begin 
				/****** Cannot delete Gold As Option in use in xTxn table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxn where xTGldAs= PMCd)
				if @row>0 
				Begin
					select @errmsg='Cannot Delete As Gold As Option Is Used In xTxn Table'
					goto ErrHndlr
				end

				/****** Cannot delete Gold As Option in use in xTxnDsg table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxnDsg where xTdGldAs= PMCd)
				if @row> 0 
				Begin
					select @errmsg= 'Cannot Delete As Gold As Option Is Used In xTxnDsg Table'
					goto ErrHndlr
				end
			End 

			if @PTyp= 'IWTFRORD' 
			Begin 
				/****** Cannot delete Weight From Ord Option in use in xTxn table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxn  where xTIWtFrOrd= PMCd)
				if @row>0 
				Begin
					select @errmsg='Cannot Delete As Weight From Order Option Is Used In xTxn Table'
					goto ErrHndlr
				end
				/****** Cannot delete Weight From Ord Option in use in xTxnDsg table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxnDsg where xTdIWtFrOrd= PMCd)
				if @row>0 
				Begin
					select @errmsg='Cannot Delete As Weight From Order Option Is Used In xTxnDsg Table'
					goto ErrHndlr
				end
			End 

			if @PTyp= 'KT' 
			Begin 
				/****** Cannot delete Karat in use in xTxnDsg table ******/				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxnDsg where xTdKt= PMCd)
				if @row>0 
				Begin	
					select @errmsg='Cannot Delete As Karat In Use In xTxnDsg Table'
					goto ErrHndlr
				end
			End 

			if @PTyp= 'LABSCD' 
			Begin 
				/****** Cannot delete Lab Sub Code in use in xTxnLab table ******/				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxnLab where xTLMCd= PMCd and xTLSCd= PSCd)
				if @row>0 
				Begin
					select @errmsg= 'Cannot Delete As Labour Sub Code In Use In xTxnLab Table'
					goto ErrHndlr
				end

				/****** Cannot delete Lab Sub Code in use in xTxnRm table ******/
				select @row=count('x') from deleted del where PMCd= 'SET' and 
					exists (select 'x' from xTxnRm where xTrSetSCd= PSCd)
				if @row>0 
				Begin
					select @errmsg= 'Cannot Delete As Setting Code In Use In xTxnRm Table'
					goto ErrHndlr
				end
			End

			if @PTyp= 'PRTCD' 
			Begin 
				/****** Cannot delete Priority Code in use in xTxn table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxn  where xTPrtCd= PMCd)
				if @row>0 
				Begin
					select @errmsg='Cannot Delete As Priority Code In Use In xTxn Table'
					goto ErrHndlr
				end
				/****** Cannot delete Priority Code in use in xTxnDsg table ******/				select @row=count('x') from deleted del where 
				exists (select 'x' from xTxnDsg where xTdPrtCd= PMCd)
				if @row>0 
				Begin
					select @errmsg='Cannot Delete As Priority Code In Use In xTxnDsg Table'
					goto ErrHndlr
				end
			End

			if @PTyp= 'YY' 
			Begin 
				/****** Cannot delete Year in use in xTxn table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxn where xTYy= PMCd)
				if @row>0 
				Begin
					select @errmsg= 'Cannot Delete As Year In Use In xTxn Table'
					goto ErrHndlr
				end
				/****** Cannot delete Year in use in xTxnDsg table ******/
				select @row=count('x') from deleted del where 
					exists (select 'x' from xTxnDsg where xTdIdYy= PMCd)
				if @row>0 
				Begin
					select @errmsg= 'Cannot Delete As Year As A Part Of Id In Use In xTxnDsg Table'
					goto ErrHndlr
				end
			
				/****** Cannot delete Year in use in xIdMst table ******/
				select @row= count('x') from deleted del where 
					exists (select 'x' from xIdMst where xImIdYy= PMCd)
				if @row> 0 
				Begin
					select @errmsg= 'Cannot Delete As Year As A Part Of Id In Use In xIdMst Table'
					goto ErrHndlr
				end
			End
		END	

	END



return
ErrHndlr:
begin
	select @errmsg='|'+@errmsg
	raiserror(@errmsg,16,1)
	rollback transaction 
		end
end


go




go




ALTER Procedure [dbo].[ForCpy_GetRmRt] 
	(@mCmCd varchar(8), @mRmCd varchar(16), @mLn1 decimal(8,3), @mLn2 decimal(8,3), 
    @mWtPerDia decimal(8,3), @mOpt As varchar(1), @mLme decimal(10,3), @mCmCtg As char(1)='C',
    @mHistCmCtg char(1) ='', @mHistCmCd varchar(8)='', @mHistYyyyMm decimal(8,0),                         
    @mCmCurCd varchar(8), @mCmLkUpRmRt varchar(8), @mPrtKey char(1), @mRt as decimal(10,3) out)
as
begin

declare @wRmRt as decimal(10,3)
Declare @wRmCtg varchar(2), @wRmSCtg varchar(5)
Declare @ctselfcocd varchar(8)
declare @ctselfcmcd varchar(8)
declare @ErrMsg varchar(500)
Declare @wRtBy varchar(2), @wRt varchar(10), @wInc decimal(10,4), @wCnd varchar(100), @wRtByForCostRt As char(1),
  @wRtChrtByPtr varchar(2), @wRtChrtByBs varchar(2), @wcndForCostRt varchar(100),
  @wCmSeekZSelfRt bit,  @wInp decimal(5,3),  @wFrLn varchar(8), @wToLn varchar(10)
declare @wCommand nvarchar(max)
  
  SELECT @ctselfcocd='ZZZ'
  SELECT @ctSelfCmCd= 'ZSELF'

  If (@mOpt <> 'C') And (@mOpt <> 'S') 
	select 0

  If (@mCmCtg <> 'C') And (@mCmCtg <> 'S') 
	select 0

  If @mCmCurCd='' 
  begin
	select @mCmCurCd = CmCurCd from CustMst where CmCtg= @mCmCtg  and CmCd= @mCmCd 
  end
  --6.1-13 seek from custmst is removed  
  --If @mCmLkUpRmRt = '' 
  --begin
  --  select @mCmLkUpRmRt = CmLkUpRmRt from CustMst where CmCtg= @mCmCtg and CmCd= @mCmCd
  --End
   
  select @wCmSeekZSelfRt = 1 where exists (Select 'x' from CustMst where CmCtg= @mCmCtg and CmCd= @mCmCd and CmSeekZSelfRt='Y')
  
  select @wRmCtg = RmCtg ,@wRmSctg=RmSctg from RmMst where RMCd=@mRmCd 
--  If @wRmCtg = 'G' Or @wRmCtg = 'P' Or @wRmCtg = 'S' Or @wRmCtg = 'L' select 0

  declare @ws_RndOff As char(1)
  select @ws_RndOff = (select case when exists((Select 'x' from Head  Join Param on PTyp='CPYRT' and PMCd=HOrdRtRndOffRmCtg Where HCoCd=@ctselfcocd 
											and hCd=@ctSelfCmCd and ','+PDesc225+',' like '%,' + @wRmCtg + ',%' )) then  0 else 3 end)
  
    
  If(@mHistCmCtg IS NULL OR @mHistCmCtg = '') Or @mHistCmCtg <> 'P' 
  begin
    select @wRtChrtByPtr = 'P'  from RmRt where RrTcTyp= 'PTR' and RrCmCtg= @mCmCtg 
                 And RrCmCd = @mCmCd and RrCtg= @wRmCtg and 
                 RrSCtg= @wRmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' 
                 and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0 
	select @wRtChrtByPtr =isnull(@wRtChrtByPtr,'')
	

    select @wRmCtg
	select @wRmSctg
	select @wRtChrtByBs = 'BS' from RmRt where RrTcTyp= 'BS' and RrCmCtg= @mCmCtg 
                 And RrCmCd = @mCmCd and RrCtg= @wRmCtg and RrSCtg= @wRmSCtg and RrSTWGrp= '' and RrCd= '' and RrDmCtg= '' 
                 and RrLsCtg= '' and RrLabMCd= '' and RrFrLn= 0 and RrToLn= 0

	
	select @wRtChrtByBs =isnull(@wRtChrtByBs ,'')
  End	

  select @wRtBy = PValue from Param where PTyp= 'RMSCTG' and PMCd= @wRmCtg and PSCd= @wRmSCtg  
  select @wRtByForCostRt = @wRtBy	
  
  If @wRtChrtByPtr <> '' 
    select @wRtBy = @wRtChrtByPtr
  Else
	If @wRtChrtByBs <> '' 
		select @wRtBy = @wRtChrtByBs
 

  IF(@mHistCmCtg IS NULL OR @mHistCmCtg = '')
    select @wRt = case when @mOpt = 'S' then 'RrSalRt' else 'RrCstRt' end
  Else
    select @wRt = case when @mOpt = 'S' then 'RhSalRt' else 'RhCstRt' end

  If @mCmCtg = 'S' And @mOpt = 'C' select @mCmCtg = 'C'


  If @mOpt ='S' 
	begin
    select @wInc = RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= @mCmCtg and 
           RrCtg= @wRmCtg and RrSCtg= @wRmSCtg  and RrCmCd= @mCmCd 
    If(@wInc IS NULL OR @wInc = 0)	
      select @wInc = RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= @mCmCtg and 
           RrCtg= @wRmCtg  and RrSCtg= ''  and RrCmCd= @mCmCd
    If(@wInc IS NULL OR @wInc = 0)	 
      select @wInc = RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= @mCmCtg and 
           RrCtg= @wRmCtg and RrSCtg= @wRmSCtg and RrCmCd= @ctSelfCmCd 
	If(@wInc IS NULL OR @wInc = 0)	 	
      select @wInc = RrSalRt from RmRt where RrTcTyp= 'INC' and RrCmCtg= @mCmCtg and            
           RrCtg= @wRmCtg and RrSCtg='' and RrCmCd= @ctSelfCmCd 
	If @wInc IS NULL 
		select @wInc = 0	
	end 
  Else
    select @wInc = 0

  If @mOpt = 'S' 
	begin
    select @wInp= RrSalRt from RmRt where RrTcTyp= 'INP' and RrCmCtg= 'P' and 
           RrCtg=@wRmCtg and RrSCtg= @wRmSCtg and RrCmCd=@mCmLkUpRmRt 
    If(@wInp IS NULL OR @wInp = 0)	
      select @wInp =RrSalRt from RmRt where RrTcTyp= 'INP' and RrCmCtg= 'P' and 
           RrCtg= @wRmCtg and RrSCtg= ''  and RrCmCd= @mCmLkUpRmRt 
	If @wInp IS NULL 
		select @wInp = 0
	end  
  Else
    select @wInp = 0

  
  IF(@mHistCmCtg IS NULL OR @mHistCmCtg = '')
	begin
		select @wFrLn = 'RrFrLn'
		select @wToLn = 'RrToLn'
	end
  Else
	begin
		select @wFrLn = 'RhFrLn'
		select @wToLn = 'RhToLn'
	End

  if @wRtBy='S' 
	begin
		If @wRmCtg = 'X'  select @mLn1 = @mLme
		select @wCnd ='  Round(' + @wFrLn+ ', 4) <= Round(' + convert(varchar,@mLn1) + ' , 3) And Round(' + @wToLn + ', 4) >= Round(' + convert(varchar,@mLn1) + ' , 3) and '    
	end
  else if @wRtBy ='W' or @wRtBy='P' 
	select @wCnd =' Round(' + @wFrLn + ', 4) <= Round(' + convert(varchar,@mWtPerDia) + ', 4) And Round(' + @wToLn + ', 4) >= Round(' + convert(varchar,@mWtPerDia) + ', 4) and '
  else if @wRtBy = 'BS'   
	select @wCnd = ' Round(' + @wFrLn + ', 4) = Round(' + convert(varchar,@mLn1) + ', 3) And Round(' + @wToLn + ', 4) = Round(' + + convert(varchar,@mLn2) + ', 3) and '
  else
	select @wCnd =' '    
	
  if @wRtByForCostRt='S'
	begin
		If @wRmCtg = 'X' select @mLn1 = @mLme
		select @wcndForCostRt = ' Round(' + @wFrLn + ', 4) <= Round(' + convert(varchar,@mLn1) + ', 3) And Round(' + @wToLn + ', 4) >= Round(' + convert(varchar,@mLn1) + ', 3) and '
	end
  else if @wRtByForCostRt in ('W', 'P')
    select @wcndForCostRt = ' Round(' + @wFrLn + ', 4) <= Round(' + convert(varchar,@mWtPerDia) + ', 4) And Round(' + @wToLn + ', 4) >= Round(' + convert(varchar,@mWtPerDia) + ', 4) and '
  else if @wRtByForCostRt = 'BS'
    select @wcndForCostRt = ' Round(' + @wFrLn + ', 4) = Round(' + convert(varchar,@mLn1) + ', 3) And Round(' + @wToLn +', 4) = Round(' + convert(varchar,@mLn2) + ', 3) and '
  else 
	select @wcndForCostRt = ' '

  Declare @wAccessoriesRmIncFac varchar(max)
  SELECT @WRTCHRTBYPTR
  SELECT @WRTCHRTBYBS 

  If (@wRtChrtByPtr = '' And @wRtChrtByBs = '' ) 
  begin	
	If(@mHistCmCtg <> '')	
	begin
		Select @wCommand='Select 
						         @wRmRt= isnull((Select ' + @wRt + ' from RmRtHist where ' +
								  'RhYyyyMm=' + convert(varchar,@mHistYyyyMm) + ' and RhTcTyp=''RM'' and ' +
								  'RhCtg=''' + @wRmCtg + ''' and RhSCtg=''' +  @wRmSCtg + '''  and  ' + 
								  'RhCd=''' + @mRmCd + ''' and '+ @wCnd +'RhCmCtg=''' + @mHistCmCtg + ''' and ' + 
								  'RhCmCd=''' + @mHistCmCd + ''' and RhCmCurCd=''' + @mCmCurCd + ''' and '+ @wRt+' <> 0),' + 					  
								  case when(@wCmSeekZSelfRt = 1 Or @mOpt = 'C') then 
								  'Convert(Decimal(14,' +case when @wInc = 0 then '3' else @ws_RndOff end + '),' +
								  'IsNull((Select ' +@wRt + ' from RmRtHist where ' + 
								  'RhYyyyMm=' + convert(varchar,@mHistYyyyMm) + ' and RhTcTyp= ''RM'' and  
								  RhCtg=''' + @wRmCtg + ''' and RhSCtg=''' + @wRmSCtg + ''' and  
								  RhCd=''' + @mRmCd + ''' and ' + @wCnd +' RhCmCtg= ''C'' and RhCmCd =''' + @ctSelfCmCd + ''' and 
								  RhCmCurCd=''' + @mCmCurCd + ''' and ' + 
								  @wRt + ' <> 0)*(1 +' + convert(varchar,@wInc) + ' / 100) ,0))' else '0' end + ')' 
	end
	Else
	begin	  
/*		select @wAccessoriesRmIncFac =										
										'(case when ''' + @wRmCtg + ''' = ''X'' then ((' + convert(varchar,@mLme) + '-RrFrLn) * ' +
										'(case when ''' + @mOpt + ''' = ''C'' then RrFixMinTol else RrFixMaxTol end)  ) else 0 end) '*/
										
		select @wAccessoriesRmIncFac =										
										'(case when ''' + @wRmCtg + ''' = ''X'' then ((' + convert(varchar,@mLme) + '-RrFrLn) * ' +
										'(case when ''' + @mOpt + ''' = ''C'' then 0 else 0  end)  ) else 0 end) '
																									         
		select @wCommand='Select @wRmRt=' + 
			     'isnull((Select ' + @wRt + '+' + @wAccessoriesRmIncFac + 
                 ' from RmRt where RrTcTyp=''RM'' and RrCtg=''' + @wRmCtg + ''' and RrSCtg=''' + @wRmSCtg + ''' and 
				 RrCd=''' + @mRmCd + ''' and ' + @wCnd + ' RrCmCtg= ''P'' and RrCmCd ='''+  @mCmLkUpRmRt + ''' and 
				 RrCmCurCd=''' + @mCmCurCd + ''' and ' +@wRt +  '<> 0)* (1 + '  + convert(varchar,@wInp) + ' / 100) ,  ' +					 				 
                 'isnull((Select ' + @wRt  + '+' + @wAccessoriesRmIncFac + 
                 ' from RmRt where RrTcTyp=''RM'' and RrCtg=''' + @wRmCtg + ''' and RrSCtg=''' + @wRmSCtg + ''' and  
                 RrCd=''' + @mRmCd + ''' and '+@wCnd + ' RrCmCtg=''' +  @mCmCtg + ''' and 
                 RrCmCd =''' + @mCmCd + ''' and RrCmCurCd=''' + @mCmCurCd + ''' and ' + @wRt + ' <> 0) ,' + 
                 ' ' + (case when @wCmSeekZSelfRt =1 Or @mOpt = 'C' 
				 then 'Convert(Decimal(14, '+ case when @wInc = 0 then '3' else @ws_RndOff end +'),' 
				 +'IsNull((Select ' + @wRt + '+' + @wAccessoriesRmIncFac +					
                 ' from RmRt where RrTcTyp= ''RM'' and RrCtg=''' + @wRmCtg + ''' and RrSCtg=''' + @wRmSCtg + ''' and 
                 RrCd=''' + @mRmCd + ''' and '+ @wCnd + ' RrCmCtg=''' + @mCmCtg + ''' and 
                 RrCmCd =''' + @ctSelfCmCd + ''' and RrCmCurCd=''' +  @mCmCurCd + ''' and '                                    
                 +@wRt + ' <> 0)*(1 +'  + convert(varchar, @wInc) + ' / 100) , 0))'  else '0' END )+ '))'
	End
  End
  Else If (@wRtChrtByPtr = 'P' Or @wRtChrtByBs = 'BS')
  begin	
    If(@mHistCmCtg <> '')
    begin
      If (@mOpt = 'C')	  
        select @wCommand='Select @wRmRt=' + @wRt + ' from RmRtHist  
                where RhYyyyMm=' + convert(varchar,@mHistYyyyMm) + ' and RhTcTyp=''RM'' and 
                RhCtg=''' + @wRmCtg + ''' and RhSCtg=''' + @wRmSCtg + ''' and 
                RhCd=''' + @mRmCd + ''' and ' + @wcndForCostRt + ' RhCmCtg= ''C'' and RhCmCd =''' + 
                @ctSelfCmCd + ''' and RhCmCurCd=''' + @mCmCurCd + ''' and ' + @wRt + ' <> 0'
      Else	  
        select @wCommand='Select @wRmRt=' + @wRt + ' from RmRtHist 
                where RhYyyyMm=' + convert(varchar,@mHistYyyyMm) + ' and RhTcTyp= ''RM'' and 
                RhCtg=''' + @wRmCtg + ''' and RhSCtg=''' + @wRmSCtg + ''' and 
                RhCd=''' + @mRmCd + ''' and ' + @wCnd + ' RhCmCtg= ''C'' and RhCmCd =''' + 
                @mCmCd + ''' and RhCmCurCd=''' + @mCmCurCd + ''' and ' + @wRt + ' <> 0'
	End
	else
  
	begin
    If (@mOpt = 'C') 
      select @wCommand='select @wRmRt=isnull((Select  '+ @wRt + ' from RmRt where RrTcTyp=''RM'' and 
				RrCtg='''+ @wRmCtg +''' and RrSCtg='''+ @wRmSCtg +''' and 
                RrCd='''+ @mRmCd +''' and '+ @wCnd + ' RrCmCtg='''+ @mCmCtg +''' and 
                RrCmCd ='''+ @mCmCd +''' and RrCmCurCd='''+ @mCmCurCd +''' and ' +@wRt + ' <> 0),
                isnull((Select ' + @wRt + ' from RmRt where RrTcTyp= ''RM'' and 
				RrCtg='''+ @wRmCtg +''' and RrSCtg=''' + @wRmSCtg +''' and 
				RrCd='''+ @mRmCd +''' and ' + @wcndForCostRt + ' RrCmCtg='''+ @mCmCtg +''' and RrCmCd ='''+ @ctSelfCmCd +''' and 
				RrCmCurCd='''+ @mCmCurCd +''' and '+ @wRt + ' <> 0),0))'
    Else				
     select @wCommand='Select @wRmRt= '+ @wRt +' from RmRt where RrTcTyp= ''RM'' and 
              RrCtg=''' + @wRmCtg + ''' and RrSCtg=''' + @wRmSCtg + ''' and 
              RrCd=''' + @mRmCd + ''' and ' + @wCnd + ' RrCmCtg=''' + @mCmCtg + ''' and RrCmCd =''' + @mCmCd + ''' and 
              RrCmCurCd=''' + @mCmCurCd + ''' and ' + @wRt + ' <> 0 '               
    End
  
  end
  select 'Command:'+isnull(@wCommand ,'')
  BEGIN TRY

  exec sp_executesql @wcommand,
						N'@mCmCd varchar(8), 
						@mRmCd varchar(16), 
                        @mLn1 decimal(8,3), 
						@mLn2 decimal(8,3), 
                        @mWtPerDia decimal(8,3), 
					    @mOpt As varchar(1),                         
					    @mLme decimal(10,3), 
						@mCmCtg As char(1),
                        @mHistCmCtg char(1) , 
						@mHistCmCd varchar(8),
                        @mHistYyyyMm decimal(8,0),                         
						@mCmCurCd varchar(8), 
                        @mCmLkUpRmRt varchar(8),                         
						@mPrtKey char(1),						
						@wRmCtg char(2),
						@wRmSCtg varchar(5),
						@wInp decimal(5,3),  
						@wRmRt decimal(10,3) output'
						,@mCmCd =@mCmCd, 
						@mRmCd =@mRmCd, 
                        @mLn1 =@mLn1, 
						@mLn2 =@mLn2, 
						@mWtPerDia =@mWtPerDia, 
						@mOpt =@mOpt,                         
						@mLme =@mLme, 
						@mCmCtg =@mCmCtg,
                        @mHistCmCtg =@mHistCmCtg, 
						@mHistCmCd =@mHistCmCd,
                        @mHistYyyyMm =@mHistYyyyMm,                         
						@mCmCurCd =@mCmCurCd, 
                        @mCmLkUpRmRt =@mCmLkuPRmRt,                         
						@mPrtKey =@mPrtKey,
						@wRmCtg=@wRmCtg,
						@wRmSCtg=@wRmSCtg,
						@wInp=@wInp,
						@wRmRt=@wRmRt output
						END TRY
						BEGIN CATCH
						 SELECT ERROR_NUMBER() AS ErrorNumber
         ,ERROR_SEVERITY() AS ErrorSeverity
         ,ERROR_STATE() AS ErrorState
         ,ERROR_PROCEDURE() AS ErrorProcedure
         ,ERROR_LINE() AS ErrorLine
         ,ERROR_MESSAGE() AS ErrorMessage
         END CATCH

Select @mrt= isnull(@wrmrt ,0)
RETURN @mrt
END


