SET TRANSACTION ISOLATION LEVEL READ COMMITTED

GOGO


Insert into TrgDisa values ('OrdMst')
Insert into TrgDisa values ('InvHd')
GOGO

Update OrdMst set OmCmCtg = 'Z', OmCmCd = OmCoCd,
OmCmCurCd = (Select CmCurCd from CustMst where CmCtg= 'Z' and
                        CmCd= OmCoCd)
where OmTc in ('JMV','JCN','JMT','LMV','LCN')

GOGO
Update InvHd set InCmCd = InCoCd 
where InTc in ('JMV','JCN','JMT','LMV','LCN')
GOGO

Delete from TrgDisa
GOGO


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
 declare @wIsJLTc char(1)
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
  select @wTcTyp= max(PValue),@wIsJLTc = max(PValue1) from inserted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''  
  if @wTcTyp = '' select @wTcTyp = 'N'
 end  
  
 if @mode in ('D')  
 begin  
  select @wTcTyp=max(PValue), @wIsJLTc = max(PValue1) from deleted join Param on PTyp= 'TC' and PMCd= OdTc and PSCd= ''   
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
    where OdPrtKey=@wCurrPrtn and OmReYy= '' And OmReChr= '' And OmReNo=0 And @wIsJLTc = 'N' and
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
   where OdPrtKey=@wCurrPrtn and OmReYy= '' And OmReChr= '' And OmReNo=0 And  @wIsJLTc = 'N' and    
    (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('PL', 'QT', 'SO', 'QB', 'JT') and     
	(Select TcMst.TmJLooseTC from TcMst where TmTc = OdTc and tmcocd = OdCoCd)='N' and
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
  (Select PMCd from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT') <> 'sri')   
  and (OdBYy <> '' or OdBChr <> '' or OdBNo <> 0)    
  if @row> 0     
  begin    
   select @errmsg= 'Bag Number Should Be Blank For This Voucher Type'    
   goto ErrHndlr    
  end    
    
  /*** Check For validity of Bag Number ***/    
  select @row= count('x') from inserted Ins where Ins.OdPrtKey=@wCurrPrtn and    
   ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('QB')   
  Or ((Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('JT') and OdTc ='sri'))  
 and not exists (Select 'x' From Bag Where BCoCd= OdCoCd and BYy= OdBYy     
      and BChr = OdBChr And BNo= OdBNo and BLoc = 'PFG'     
      and BOdDmCd= OdDmCd and BPrtKey=OdPrtKey)    
   And Not Exists (Select 'x' From OrdDsg Od Where Od.OdCoCd=ins.OdCoCd And Od.OdTc In ('QB', 'sri') And     
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
   (Select PValue from Param where PTyp= 'TC' and PMCd= OdTc) in ('JT') and OdTc ='sri' and  
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
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and not exists (select 'x' from vParam where vPCoCd= OdCoCd and     
     vPTyp= 'PRDSEQ' and vPMCd= OdPrdSeq )    
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
select @row= count('x') from inserted where OdPrtKey=@wCurrPrtn and @wIsJLTc = 'N' and    
   not exists (select 'x' from Param where PTyp= 'PRTCD' and PMCd= OdPrtCd )    
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
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'KT' and PMCd= OdKt)    
  if @row> 0     
  begin    
   select @errmsg= 'Karatage Not Defined'    
   goto ErrHndlr    
  end    
    
  /******  OdMulBy  >  0  ******/    
  /* Current Partition */     
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdMulBy<= 0     
  if @row> 0     
  begin    
   select @errmsg= 'Multiplying Factor Cannot Be <= Zero'    
   goto ErrHndlr    
  end    
   /******  OdFixPrc  =  PMCd ('YN')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
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
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPRtn and OdGldAs<> '' and     
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
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'GLDAS' and PMCd= OdLabAs)    
  if @row> 0     
  begin    
   select @errmsg= 'Labour As Value Not Defined'    
   goto ErrHndlr    
  end    
    
  /******  OdIWtEqOrd  =  PMCd ('YN')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'YN' and PMCd= OdIWtEqOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. Equal To Order Wt. (Y/N)'    
   goto ErrHndlr    
  end    
    
  /******  If OdIWtEqOrd  =  'Y'  then  OdIWtFrOrd  =  Blank      
   If OdIWtEqOrd  =  'N'  then  OdIWtFrOrd  =  Blank  or  PMCd ('IWTFRORD')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdIWtFrOrd<> '' and OdPrtKEy=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'IWTFRORD' and PMCd= OdIWtFrOrd)    
  if @row> 0     
  begin    
   select @errmsg= 'Invoice Wt. From Order Not Defined'    
   goto ErrHndlr    
  end     
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdIWtFrOrd<> '' and OdIWtEqOrd= 'Y')     
  if @row> 0     
  begin    
   select @errmsg= 'Mismatch in values InvWtEqOrd and InvWtFrOrd'    
   goto ErrHndlr    
  end     
    
  /******  OdIGldRtEqOrd  =  PMCd ('YN')  ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
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
  select @row= count('x') from inserted Ins where @wIsJLTc = 'N' and Ins.OdPrtKEy=@wCurrPRtn and    
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
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPrtn and    
   not exists (select 'x' from Param where PTyp= 'WH' and PMCd= OdWh)    
  if @row> 0      
  begin    
   select @errmsg= 'W/H Set Option Not Defined'    
   goto ErrHndlr    
  end    
    
  /******  OdLine  =  vPMCd ('LINE')  ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
   not exists (select 'x' from vParam where vPCoCd= OdCoCd and vPTyp= 'LINE' and vPMCd= OdLine and vPMCd<> 'ZZZ')    
  if @row> 0    begin    
   select @errmsg= 'Line Not Defined/ Line Cannot Be ''ZZZ'' '    
   goto ErrHndlr    
  end    
    
  /******  OdWDiaAvlblDt And OdHDiaAvlblDt <>  '01/01/1980'  or  '01/01/1900'  (The dates '01/01/1980'  and  '01/01/1900'  are considered as blank) ******/    
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and (OdWDiaAvlblDt in('01/01/1980', '01/01/1900') or     
   OdHDiaAvlblDt in('01/01/1980', '01/01/1900'))    
  if @row> 0     
  begin    
   select @errmsg= 'Wax/ Hand Set Diamond Availability Date Cannot Be Blank'    
   goto ErrHndlr    
  end    
    
  /****** OdGmChk = PMCd('CPYRT') ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPRtKey=@wCurrPRtn and    
    not exists (select 'x' from Param where PTyp= 'CPYRT' and PMCd= OdGmChk)    
  if @row> 0     
  begin    
   select @errmsg= 'Gram Check Option Not Defined'    
   goto ErrHndlr    
  end    
    
  /****** OdChkTol = PMCd('YN') ******/     
  /* Current Partition */    
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
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
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdDelDt= '01/01/1900' or (OdDelDt<> '01/01/1980' and     
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
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdExpDelDt= '01/01/1900' or (OdExpDelDt<> '01/01/1980' and     
   OdExpDelDt < (Select OmDt from OrdMst where OmIdNo=OdOmIdNo and OmCoCd= OdCoCd     
    and OmTc= OdTc and OmYy= OdYy and OmChr= OdChr and OmNo= OdNo and OmPrtKey=OdPrtKey))     
    
/* **** Zubin 211 **** */     
  if @row> 0     
  begin    
   select @errmsg= 'Export Delivery Date Cannot Be Less Than Order Date'    
   goto ErrHndlr   end    
  /****** Order Qty, Prd Qty should be >= 0 ******/    
    
  /* Current Partition */    
  select @row= count('x') from inserted where oDtC <> 'JCN' and @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdOrdQty < 0 or OdPrdQty< 0    
  if @row> 0     
  begin    select @errmsg= 'Order/ Prd Qty Should Be >= 0 '    
   goto ErrHndlr    
  end    
    
  /****** Order Quantity should be >= Export Quantity ******/    
  /* Current Partition */    
 select @row= count('x') from inserted where odtc <> 'JCN' and @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and    
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
  select @row= count('x') from inserted where @wIsJLTc = 'N' and OdPrtKey=@wCurrPrtn and OdBagPcs<= 0     
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
    exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='sri')  
   If @Row > 0  
   Begin  
    Select @ErrMsg= 'A Transaction of Later Date Exists For this Bag'  
    GoTo ErrHndlr  
   end  
      
	
   /* Manali Trading Module */  
   Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo and Om.OmPrtKey=Ins.OdPrtKey     
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       (   
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo and Om1.OmPrtKey=Od.OdPrtKey    
            Where Od.OdPrtKey=Ins.OdPrtKey and Od.OdCoCd=(Case when ins.OdTc in ('JIN','LIN')  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
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
  
   /* Manali Trading Module */  
   Select @Row=Count('x') from Inserted Ins  
    Join OrdMst Om On Om.OmCoCd=Ins.OdCoCd and Om.OmTc=Ins.OdTc and Om.OmChr=Ins.OdChr and Om.OmNo=Ins.OdNo and Om.OmIdNo=Ins.OdOmIdNo and Om.OmPrtKey=Ins.OdPrtKey     
    Where exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmBaseTcs<>'' and   
       (   
       (TmBaseMustYN='Y' and (Ins.OdBaseIdKey=0  Or    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo and Om1.OmPrtKey=Od.OdPrtKey    
            Where Od.OdPrtKey=Ins.OdPrtKey and Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
            and Od.OdJLRmCd=Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1 =Ins.OdJLLn1 and
			Od.OdJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.OdJLLn3    
            and (Od.OdDmStkNo=Ins.OdDmStkNo/*Or (Od.OdDmStkNo=0 Or Ins.OdDmStkNo=0)*/)  
            and Om1.OmCmCtg=Om.OmCmCtg and Om1.OmCmCd=Om.OmCmCd  
            and Od.OdIdNo=Ins.OdBaseIdKey), '')))   
       Or  
       (TmBaseMustYN='N' and Ins.OdBaseIdKey<>0  AND    
       TmBaseTcs<>IsNull((Select Od.OdTc from OrdDsg Od   
            Join OrdMst Om1 On Om1.OmCoCd=Od.OdCoCd and Om1.OmTc=Od.OdTc and Om1.OmChr=Od.OdChr and Om1.OmNo=Od.OdNo and Om1.OmIdNo=Od.OdOmIdNo and Om1.OmPrtKey=Od.OdPrtKey    
            Where Od.OdPrtKey=Ins.OdPrtKey and  Od.OdCoCd=(Case when ins.OdTc='LIN'  Then Om.OmBaseCoCd Else Ins.OdCoCd End)  
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

   Select @Row=Count('x') from Inserted Ins Where OdPrtKey=@wCurrPrtn  
    And exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'N'
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
  
   Select @Row=Count('x') from Inserted Ins Where OdPrtKey=@wCurrPrtn  
    And exists(Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  and @wIsJLTc = 'Y'
    And exists(Select 'x' from TcMst Where TmCoCd=Ins.OdCoCd and TmTc=Ins.OdTc and TmRefTcs<>'' and   
         Ins.OdRefIdKey<>0 And  
         ','+TmRefTcs+',' Not Like   
          '%,'+IsNull((Select OdTc from OrdDsg Od  
                Where OdPrtKey=Ins.OdPrtKey and OdCoCd=Ins.OdCoCd and OdIdNo=Ins.OdRefIdKey   
                and Od.ODJLRmCd =Ins.OdJLRmCd and Od.OdJLLotNo=Ins.OdJLLotNo and Od.OdJLLn1=Ins.OdJLLn1 
				and Od.ODJLLn2 = Ins.ODJLLn2 and Od.OdJLLn3 = Ins.ODJLLn3 
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
    exists(Select 'x' from OrdDsg Od Where  (Del.OdTc in ('JIN','LIN') Or Od.OdCoCd=Del.OdCoCd) and Od.OdPrtKey=Del.OdPrtKey and   
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
   Where OdPrtKey=@wCurrPrtn and @wIsJLTc = 'N' and
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
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'N' 
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

  Begin  
  Select @row=Count('x') from deleted Del   
   Join TcMst On TmCoCd=Del.OdCoCd and TmTc=Del.OdTc  
   Where OdPrtKey=@wCurrPrtn and @wIsJLTc = 'Y' and
   exists(Select 'x' from OrdDsg Od Where  (Del.OdTc='LIN' Or  Od.OdCoCd=Del.OdCoCd) and Od.OdPrtKey=Del.OdPrtKey and   
      Od.OdBaseIdKey=Del.OdIdNo)  
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
    and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
    and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr   
    and ins.OdJLRmCd=del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJLLn2 = del.odJLLn2 and ins.odJLLn3 = del.odJLLn3 
    /*((TmBaseMustYn='Y' and ins.OdDmId=del.OdDmId) or (TmBaseMustYn='N' and (ins.OdDmId=0 or del.OdDmId=0) ))*/  
      )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Base Voucher'  
   goto ErrHndlr  
  End  

    
  Select @row=Count('x') from deleted Del Where OdPrtKey=@wCurrPrtn and  
   exists(Select 'x' from OrdDsg Od Where Od.OdCoCd=Del.OdCoCd and Od.OdPrtKey=Del.OdPrtKey and   
      Od.OdRefIdKey=Del.OdIdNo)  and @wIsJLTc = 'Y' 
   and not exists (Select 'x' from inserted ins where ins.OdIdNo=del.OdIdNo and ins.OdCoCd= del.OdCoCd     
     and ins.OdTc= del.OdTc and ins.OdYy= del.OdYy and ins.OdChr= del.OdChr     
     and ins.OdNo= del.OdNo and ins.OdSr= del.OdSr   
     and ins.OdJLRmCd =del.OdJLRmCd and ins.OdJLLotNo=del.OdJLLotNo and ins.OdJLLn1 =del.OdJLLn1 and ins.OdJlLn2 = del.OdJlLn2 and ins.OdJLLn3 = del.ODJLLn3
     )    
  If @row > 0  
  Begin  
   Select @errmsg='Cannot Change RmCd/LotNo/Size as Order Serial Is Used as Reference Voucher'  
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
  @wInFob= convert (decimal (16, 2), InFOB- IsNull((Select sum((del.OdSalPrc*del.OdOrdQty)+del.OdJLVchVal) from deleted del where InPrtKey=del.OdPrtKey and del.OdCoCd= InCoCd and del.OdTc= InTc and del.OdYy= InYy   
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
  if @wIsJLTc = 'N'
  begin
  
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
   end
   else
   begin
  update DmYyMm set  
    DDmCrWt= DDmCrWt- (select isnull(sum(Del.OdJLWt), 0) from Deleted Del  
   Join OrdMst on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.ODNo where OmFrDmDc='C' and Del.OdPrtKey=DPrtKey and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))= DYyMm),  
  DDmDrWt= DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del  
   Join OrdMst on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr  
   and OmNo= Del.OdNo where OmFrDmDc='D' and Del.OdPrtKey=DPrtKey and Del.OdCoCd= DCoCd and Del.OdJLRmCd= DJLRmCd 
   and Del.OdJLLotNo= DJLLotNo and Del.OdJLLn1= DJLLn1 and Del.OdJLLn2= DJLLn2 and Del.OdJLLn3 = DJLLn3 and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmPrtKey=Deleted.OdPrtKey and OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc and OmYy= Deleted.OdYy  
   and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdPrtKey=DPrtKey and Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd= DJLRmCd 
   and Deleted.OdJLLotNo= DJLLotNo and Deleted.OdJLLn1= DJLLn1 and Deleted.OdJLLn2= DJLLn2 and Deleted.OdJLLn3 = DJLLn3 and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
   end
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
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri' ),  
  SRmCrWt= SRmCrWt-(select isnull(sum(FrRmWt *(Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmFrRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmDrQty= SRmDrQty-(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0) 
  
   From FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmFrRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmDrWt= SRmDrWt-(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmFrRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri')   
  from FgRm   
  Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr    
  Join Deleted On Deleted.OdPrtKey=FdPrtKey and Deleted.OdCoCd=FdCoCd and Deleted.OdBYy=FdBYy and Deleted.OdBChr=FdBChr and Deleted.OdBNo=FdBNo  
  Join OrdMst On OmPrtKey=Deleted.OdPrtKey and OmCoCd=Deleted.OdCoCd and OmTc=Deleted.OdTc and OmYy=Deleted.OdYy and OmChr=Deleted.OdChr and OmNo=Deleted.OdNo  
  where FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd= SRmCd and FrLotNo= SLotNo   
  and FrRmSz2= SRmSz2   and FrRmSz3= SRmSz3   and FrRmSz= SRmSz and FrRmStkRt= SRmStkRt   
  and OmFrRmLoc= SLoc and convert(int,convert(varchar(6),OmDt,112))=SYyMm  
  and (OmFrRmDc='D'  or OmFrRmDc='C')  
  and OdTc= 'sri'  
     
  if @mode = 'D'    Begin  
   /******* (BCls= 'N') Bag is again opened for transaction ******/  
   update Bag set  
   BCls= 'N', BClsDt= '01/01/80', BCnvJtYn='N'  
   from deleted where BPrtKey=@wCurrPrtn and BCoCd= OdCoCd and BYy= OdBYy and BChr= OdBChr and BNo= OdBNo    
   and OdTc= 'sri'   
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
  @wInFob = convert (decimal (16, 2), InFOB+ IsNull((Select sum((ins.OdSalPrc*ins.OdOrdQty)+ins.OdJLVchVal) from inserted ins where   
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
  

  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr, DDmStkNo, DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt)  
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy,OdDmStkChr,OdDmStkNo,OmFrDmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0,@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0  
   from inserted join OrdMst on OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DPrtKey=OdPrtKey and DCoCd= OdCoCd and DDmCd=OdDmCd   
        and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy= DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmFrDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and
        DYyMm= convert(int,convert(varchar(6),OmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
         
  /* Current Partition */   
  if @wIsJLTc ='N'
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
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='C' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd 
   and OmFrDmLoc=DLoc  and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmFrDmDc='D'  and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd= DJLRmCd 
   and Ins.OdJLLotNo= DJLLotNo and Ins.OdJLLn1= DJLLn1 and Ins.OdJLLn2= DJLLn2 and Ins.OdJLLn3 = DJLLn3 
   and OmFrDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmPrtKey=Inserted.OdPrtKey and OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdPrtKey=DPrtKey and Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd= DJLRmCd 
   and Inserted.OdJLLotNo= DJLLotNo and Inserted.OdJLLn1= DJLLn1 and Inserted.OdJLLn2= DJLLn2 and Inserted.OdJLLn3 = DJLLn3  and OmFrDmLoc= DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm  
   and (OmFrDmDc='D' or OmFrDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 end  
  
 /*If exists (Select 'x' from Inserted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='JCN'))*/  
 If @wTcTyp = 'JT'   
 Begin  
  Insert into SYyMm(SPrtKey,SCoCd,SRmCd,SLotNo,SRmSz,SRmSz2,SRmSz3,SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)   
   select distinct FrPrtKey,FrCoCd,FrRmCd,FrLotNo,FrRmSz,FrRmSz2,FrRmSz3,FrRmStkRt,OmFrRmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0,0,0  
   From FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted On OdPrtKey=FdPrtKey and OdCoCd=FdCoCd and OdBYy=FdBYy and OdBChr=FdBChr and OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=OdPrtKey and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  SYyMm where SPrtKey=FrPrtKey and SCoCd= FrCoCd and SRmCd=FrRmCd and SLotNo=FrLotNo   
   and SRmSZ=FrRmSz and SRmSZ2=FrRmSz2 and SRmSZ3=FrRmSz3 and SRmStkRt=FrRmStkRt and SLoc=OmFrRmLoc and  
   SYyMm= convert(int,convert(varchar(6),OmDt,112))) and OdTc= 'sri'   
     
  
  /* Current Partition */  
  Update SYyMm set  
  SRmCrQty= SRmCrQty+(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0) 
  
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmFrRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and SRmSZ2=FrRmSz2 and SRmSZ3=FrRmSz3 and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmCrWt=SRmCrWt+(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmFrRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmDrQty=SRmDrQty+(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)  
 
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmFrRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmDrWt=SRmDrWt+(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmFrRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmFrRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri')   
  from FgRm   
  Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr    
  Join Inserted On Inserted.OdPrtKey=FdPrtKey and Inserted.OdCoCd=FdCoCd and Inserted.OdBYy=FdBYy and Inserted.OdBChr=FdBChr and Inserted.OdBNo=FdBNo  
  Join OrdMst On OmPrtKey=Inserted.OdPrtKey and OmCoCd=Inserted.OdCoCd and OmTc=Inserted.OdTc and OmYy=Inserted.OdYy and OmChr=Inserted.OdChr and OmNo=Inserted.OdNo  
  where FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd= SRmCd and FrLotNo= SLotNo   
  and FrRmSz= SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt= SRmStkRt   
  and OmFrRmLoc= SLoc and convert(int,convert(varchar(6),OmDt,112))=SYyMm  
  and (OmFrRmDc='D'  or OmFrRmDc='C')  
  and OdTc= 'sri'  
    
  if @mode in ('A')  
  Begin   
   update Bag set   
   BCls= 'Y',   
   BCnvJtYn ='Y',  
   BClsDt= (Select OmDt From OrdMst Where OmCoCd= OdCoCd And OmTc = OdTc   
     And OmYy = OdYy And OmChr= OdChr And OmNo= OdNo And OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo)   
   from inserted where BCoCd= OdCoCd and BYy= OdBYy and BChr= OdBChr and BNo= OdBNo and OdTc= 'sri'    
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
 if @wIsJLTc = 'N'
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
  else
  update DmYyMm set  
  DDmCrWt=DDmCrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst  
   on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='C' and Del.OdPrtKey=DPrtKey and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt-(select isnull(sum(Del.OdJLWt),0) from Deleted Del Join OrdMst   
   on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   where OmToDmDc='D' and Del.OdPrtKey=DPrtKey and Del.OdCoCd= DCoCd and Del.OdJLRmCd = DJLRmCd and Del.OdJLLotNo = DJLLotNo and Del.OdJLLn1 = DJLLn1 and
   Del.OdJlLn2 = DJLLn2 and Del.OdJllN3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from Deleted Join OrdMst on OmPrtKey=Deleted.OdPrtKey and OmIdNo=Deleted.OdOmIdNo and OmCoCd= Deleted.OdCoCd and OmTc= Deleted.OdTc  
   and OmYy= Deleted.OdYy and OmChr= Deleted.OdChr and OmNo= Deleted.OdNo  
   where Deleted.OdPrtKey=DPrtKey and Deleted.OdCoCd= DCoCd and Deleted.OdJLRmCd = DJLRmCd and Deleted.OdJLLotNo = DJLLotNo and Deleted.OdJLLn1 = DJLLn1 and
   Deleted.OdJlLn2 = DJLLn2 and Deleted.OdJllN3 = DJLLn3  and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))= DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 /*If exists (Select 'x' from Deleted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='sri'))*/  
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
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmCrWt= SRmCrWt-(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmToRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmDrQty= SRmDrQty-(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0) 
  
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmToRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmDrWt= SRmDrWt-(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Deleted Del On Del.OdPrtKey=FdPrtKey and Del.OdCoCd=FdCoCd and Del.OdBYy=FdBYy and Del.OdBChr=FdBChr and Del.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Del.OdPrtKey and OmCoCd=Del.OdCoCd and OmTc=Del.OdTc and OmYy=Del.OdYy and OmChr=Del.OdChr and OmNo=Del.OdNo  
   where OmToRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri')   
  from FgRm   
  Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr    
  Join Deleted On Deleted.OdPrtKey=FdPrtKey and Deleted.OdCoCd=FdCoCd and Deleted.OdBYy=FdBYy and Deleted.OdBChr=FdBChr and Deleted.OdBNo=FdBNo  
  Join OrdMst On OmPrtKey=Deleted.OdPrtKey and OmCoCd=Deleted.OdCoCd and OmTc=Deleted.OdTc and OmYy=Deleted.OdYy and OmChr=Deleted.OdChr and OmNo=Deleted.OdNo  
  where FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd= SRmCd and FrLotNo= SLotNo   
  and FrRmSz= SRmSz and FrRmSz2= SRmSz2 and FrRmSz3= SRmSz3 and FrRmStkRt= SRmStkRt   
  and OmToRmLoc= SLoc and convert(int,convert(varchar(6),OmDt,112))=SYyMm  
  and (OmToRmDc='D'  or OmToRmDc='C')  
  and OdTc= 'sri'  
 End  
  
 /*If exists (Select 'x' from Deleted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT')  
   and Exists (Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTC and TmBaseTCs<>''))*/  
 If @wTcTyp = 'JT'     
 Begin   
  Update OrdDsg   
  Set OrdDsg.OdExpQty= OrdDsg.OdExpQty- IsNull((Select sum(del.OdOrdQty) from deleted del where   
   del.OdPrtKey=OrdDsg.OdPrtKey and  (del.OdTc='JIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt- IsNull((Select sum(del.OdJLWt) from deleted del where   
   del.OdPrtKey=OrdDsg.OdPrtKey and  (del.OdTc='LIN' Or del.OdCoCd= OrdDsg.OdCoCd) and del.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from deleted where Deleted.OdPRtKey=@wCurrPrtn and OrdDsg.OdPrtKey=Deleted.OdPrtKey and  
    OrdDsg.OdCoCd= deleted.OdCoCd and OrdDsg.OdIdNo=deleted.OdBaseIdkey  
   and Exists (Select 'x' from TcMst Where TmCoCd=deleted.OdCoCd and TmTc=deleted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls = (case when @wIsJLTc ='N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
				else (case when OrdDsg.OdJLWt= OrdDsg.OdJLExpWt and OrdDsg.OdJLWt<> 0 then 'Y' else 'N' end)   end)
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
  insert into DmYyMm(DPrtKey,DCoCd,DDmCd,DOdSfx,DDmSz,DDmStkYy,DDmStkChr,DDmStkNo,DLoc,DYyMm,DDmDrQty,DDmCrQty,DJLTag,DJLRmCd,DJLLotNo,DJLLn1,DJLLn2,DJLLn3,DDmDrWt,DDmCrWt)
   select distinct OdPrtKey, OdCoCd, OdDmCd, OdSfx, OdDmSz,OdDmStkYy, OdDmStkChr, OdDmStkNo,OmToDmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0,
					@wIsJLTc,OdJLRmCd,ODJLLotNo,OdJLLn1,ODJLLn2,ODJLLn3,0,0  
     
   from inserted join OrdMst on OmPrtKey=OdPrtKey and OmIdNo=OdOmIdNo and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  DmYyMm where DPrtKey=OdPrtKey and DCoCd= OdCoCd and DDmCd=OdDmCd   
       and OdSfx= DOdSfx and OdDmSz= DDmSz and OdDmStkYy = DDmStkYy and OdDmStkChr = DDmStkChr and OdDmStkNo = DDmStkNo and DLoc=OmToDmLoc and  
			OdJLRmCd= DJLRmCd and OdJLLotNo = DJLLotNo and ODJlLn1 = DJLLn1 and OdJLLn2 = DJLLn2 and OdJLLn3 = DJLLn3 and
       DYyMm= convert(int,convert(varchar(6),OmDt,112)))  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  
  if @wIsJLTc = 'N'
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
else
  update DmYyMm set  
  DDmCrWt=DDmCrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='C' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm),  
  DDmDrWt=DDmDrWt+(select isnull(sum(Ins.OdJLWt),0) from Inserted Ins Join OrdMst   
   on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   where OmToDmDc='D' and Ins.OdPrtKey=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.ODJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3 = DJLLn3 
   and OmToDmLoc=DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm)  
  from Inserted Join OrdMst on OmPrtKey=Inserted.OdPrtKey and OmIdNo=Inserted.OdOmIdNo and OmCoCd= Inserted.OdCoCd and OmTc= Inserted.OdTc  
   and OmYy= Inserted.OdYy and OmChr= Inserted.OdChr and OmNo= Inserted.OdNo  
   where Inserted.OdPrtKey=DPrtKey and Inserted.OdCoCd= DCoCd and Inserted.OdJLRmCd = DJLRmCd and Inserted.ODJLLotNo = DJLLotNo and Inserted.ODJLLn1 = DJLLn1 and Inserted.ODJLLn2 = DJLLn2 and 
   Inserted.ODJLLn3 = DJLLn3 
   and OmToDmLoc= DLoc  
   and convert(int,convert(varchar(6),OmDt,112))=DYyMm  
   and (OmToDmDc='D' or OmToDmDc='C')  
   and OmFrDmLoc <>'' and OmToDmLoc<>''  
 End  
  
 
 /*If exists (Select 'x' from Inserted   
   where exists (Select 'x' from Param where PTyp= 'TC' and PMCd= OdTc and PValue='JT' and PMCd='sri'))*/  
 If @wTcTyp = 'JT'     and   @wIsJLTc = 'N'

 Begin   
   insert into SYyMm(SPrtKey,SCoCd,SRmCd,SLotNo,SRmSz,SRmSz2,SRmSz3,SRmStkRt,SLoc,SYyMm,SRmDrqty,SRmDrWt,SRmCrQty,SRmCrWt)   
   select distinct FrPrtKey,FrCoCd,FrRmCd,FrLotNo,FrRmSz,FrRmSz2,FrRmSz3,FrRmStkRt,OmToRmLoc,convert(int,convert(varchar(6),OmDt,112)),0,0,0,0  
   From FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted On OdPrtKey=FdPrtKey and OdCoCd=FdCoCd and OdBYy=FdBYy and OdBChr=FdBChr and OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=OdPrtKey and OmCoCd=OdCoCd and OmTc=OdTc and OmYy=OdYy and OmChr=OdChr and OmNo=OdNo  
   where not exists(select 'x' from  SYyMm where SPrtKey=FrPrtKey and SCoCd= FrCoCd and SRmCd=FrRmCd and SLotNo=FrLotNo   
   and SRmSZ=FrRmSz and SRmSZ2=FrRmSz2 and SRmSZ3=FrRmSz3 and SRmStkRt=FrRmStkRt and SLoc=OmToRmLoc and  
   SYyMm= convert(int,convert(varchar(6),OmDt,112))) and OdTc= 'sri'   
  update SYyMm set  
  SRmCrQty= SRmCrQty+(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0) 
  
   from FgRm  
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmToRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmCrWt=SRmCrWt+(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmToRmDc= 'C' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmDrQty=SRmDrQty+(select isnull(sum(FrRmQty * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)  
 
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmToRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri') ,  
  SRmDrWt=SRmDrWt+(select isnull(sum(FrRmWt * (Case When FrFrRmLoc='FG' and FrFrRmDc= 'D' Then 1 When FrFrRmLoc='FG' and FrFrRmDc= 'C' Then -1 When FrToRmLoc='FG' and FrToRmDc= 'D' Then 1 When FrToRmLoc='FG' and FrToRmDc= 'C' Then -1 Else 0 End)), 0)   
   from FgRm   
   Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr   
   and FdNo= FrNo and FdSr=FrSr    
   Join Inserted Ins On Ins.OdPrtKey=FdPrtKey and Ins.OdCoCd=FdCoCd and Ins.OdBYy=FdBYy and Ins.OdBChr=FdBChr and Ins.OdBNo=FdBNo  
   Join OrdMst On OmPrtKey=Ins.OdPrtKey and OmCoCd=Ins.OdCoCd and OmTc=Ins.OdTc and OmYy=Ins.OdYy and OmChr=Ins.OdChr and OmNo=Ins.OdNo  
   where OmToRmDc= 'D' and FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd=SRmCd   
   and FrLotNo=SLotNo and FrRmSz=SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt=SRmStkRt and OmToRmLoc=SLoc   
   and convert(int,convert(varchar(6),OmDt,112))= SYyMm and OdTc= 'sri')   
  from FgRm   
  Join Fgd on FdPrtKey=FrPrtKey and FdCoCd= FrCoCd and FdTc= FrTc and FdYy= FrYy and FdChr= FrChr and FdNo= FrNo and FdSr=FrSr    
  Join Inserted On Inserted.OdPrtKey=FdPrtKey and Inserted.OdCoCd=FdCoCd and Inserted.OdBYy=FdBYy and Inserted.OdBChr=FdBChr and Inserted.OdBNo=FdBNo  
  Join OrdMst On OmPrtKey=Inserted.OdPrtKey and OmCoCd=Inserted.OdCoCd and OmTc=Inserted.OdTc and OmYy=Inserted.OdYy and OmChr=Inserted.OdChr and OmNo=Inserted.OdNo  
  where FrPrtKey=SPrtKey and FrCoCd= SCoCd and FrRmCd= SRmCd and FrLotNo= SLotNo   
  and FrRmSz= SRmSz and FrRmSz2=SRmSz2 and FrRmSz3=SRmSz3 and FrRmStkRt= SRmStkRt   
  and OmToRmLoc= SLoc and convert(int,convert(varchar(6),OmDt,112))=SYyMm  
  and (OmToRmDc='D'  or OmToRmDc='C')  
  and OdTc= 'sri'  
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
   Ins.OdPrtKey=OrdDsg.OdPrtKey and  (ins.OdTc='JIN' Or Ins.OdCoCd= OrdDsg.OdCoCd)  and Ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  ,
	OrdDsg.OdJLExpWt= OrdDsg.OdJLExpWt+ IsNull((Select sum(ins.OdJLWt) from Inserted Ins where   
   ins.OdPrtKey=OrdDsg.OdPrtKey and  (ins.OdTc='LIN' Or ins.OdCoCd= OrdDsg.OdCoCd) and ins.OdBaseIdKey= OrdDsg.OdIdNo), 0)  
   from Inserted where Inserted.OdPRtKey=@wCurrPrtn and OrdDsg.OdPrtKey=Inserted.OdPrtKey and  
    OrdDsg.OdCoCd= Inserted.OdCoCd and OrdDsg.OdIdNo=Inserted.OdBaseIdkey  
    and Exists (Select 'x' from TcMst Where TmCoCd=Inserted.OdCoCd and TmTc=Inserted.OdTC and TmBaseTCs<>'')  
  
  Update OrdDsg   
  Set OdCls =(case when @wIsJLTc = 'N' then (case when OrdDsg.OdOrdQty= OrdDsg.OdExpQty and Round(OrdDsg.OdOrdQty, 1)<> 0 then 'Y' else 'N' end)   
			else (case when OrdDsg.OdJLWt = OrdDsg.ODJLExpWt and OrdDsg.ODJLWt <> 0 then 'Y' else 'N' end) end)
   from Inserted Ins where OrdDsg.OdPRtKEy=Ins.OdPRtKey and (ins.OdTc = 'LIN' or ins.OdTc='JIN' Or  OrdDsg.OdCoCd= Ins.OdCoCd)   
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
if @wIsJLTc = 'N'
begin
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
else
begin
 Select @Row=Count('x') from Inserted Ins   
   /*Join OrdDsg B On B.OdPrtKey=Ins.OdPrtKey and B.OdCoCd= Ins.OdCoCd and B.OdIdNo=Ins.OdBaseIdkey */  
   where Ins.OdPRtKey=@wCurrPrtn   
    AND exists(Select 'x' from OrdDsg B Where  
      B.OdPrtKey=Ins.OdPrtKey and (Ins.OdTc='LIN' Or  B.OdCoCd= Ins.OdCoCd) and B.OdIdNo=Ins.OdBaseIdkey and B.OdJLWt < B.OdJLExpWt)  
    and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmbaseTcs<>'')  
  If @Row> 0  
  begin   
   select @errmsg= 'Order Wt Should Be >= Matched Wt for Base Voucher'    
   goto ErrHndlr    
  end
 end   
end    
  
/*if @mode in ('C','A') and exists (Select 'x' from Inserted Ins  
   Join Param On PTyp= 'TC' and (PValue = 'FG' Or (PValue='JT' and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmLocTypFr<>'' and TmLocTypTo<>'' )))   
   and PMCd= OdTc)*/  
If @mode in ('C','A') and  (@wTcTyp = 'FG' or @wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') From Inserted Ins   
   Join OrdMst on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Ins.OdPrtKEy=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdDmCd= DDmCd and Ins.OdSfx= DOdSfx and Ins.OdDmSz= DDmSz and   
     (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
    /*Ins.OdDmId= DDmId and DLoc=LocCd*/  
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
   Join OrdMst on OmPrtKey=Ins.OdPrtKey and OmIdNo=Ins.OdOmIdNo and OmCoCd= Ins.OdCoCd and OmTc= Ins.OdTc  
       and OmYy= Ins.OdYy and OmChr= Ins.OdChr and OmNo= Ins.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Ins.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End) and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
      /*LocCd in (OmFrDmLoc,OmToDmLoc)*/  
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Ins.ModUsr and UaMnuCd = Ins.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrWt)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Ins.OdPrtKEy=DPrtKey and Ins.OdCoCd= DCoCd and Ins.OdJLRmCd = DJLRmCd and Ins.OdJLLotNo = DJLLotNo and Ins.ODJLLn1 = DJLLn1 and Ins.ODJLLn2 = DJLLn2 and Ins.ODJLLn3=DJLLn3 and
	(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
    /*Ins.OdDmId= DDmId and DLoc=LocCd*/  
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''   
  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Size'  
     GoTo ErrHndlr  
     end  

end
end  
  
  
  
/*if @mode in ('D') and exists (Select 'x' from Deleted   
     Join Param On PTyp= 'TC' and (PValue = 'FG' Or (PValue='JT' and exists(Select 'x' from TcMst Where TmCoCd=OdCoCd and TmTc=OdTc and TmLocTypFr<>'' and TmLocTypTo<>'' )))   
     and PMCd= OdTc)*/  
If @mode in ('D') and  (@wTcTyp = 'FG' or @wTcTyp = 'JT')  
Begin   
if @wIsJLTc = 'N'
begin
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Del.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End)  and   
            LocCd in (case when OmTc<>'JCN' or ODORDQTY>0 then (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
      /*LocCd in (OmFrDmLoc,OmToDmLoc)*/  
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrQty)-SUM(DDmCrQty) As Decimal(10,3)) from DmYyMm  
    where Del.OdPrtKEy=DPrtKey and Del.OdCoCd= DCoCd and Del.OdDmCd= DDmCd and Del.OdSfx= DOdSfx and Del.OdDmSz= DDmSz and   
    (case when OmTc<>'JCN' OR ODORDQTY>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
     /*Del.OdDmId= DDmId and DLoc=LocCd*/  
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
 Select @Row=Count('x') From Deleted Del   
   Join OrdMst on OmPrtKey=Del.OdPrtKey and OmIdNo=Del.OdOmIdNo and OmCoCd= Del.OdCoCd and OmTc= Del.OdTc  
       and OmYy= Del.OdYy and OmChr= Del.OdChr and OmNo= Del.OdNo  
   Join Param On PTyp='TC' and PMCd=OdTc  
   Join Head On HCoCd=OdCoCd And HNegStkYn= 'N'  
   Join Loc On LocCoCd= Del.OdCoCd and LocTyp =(Case PValue When 'FG' Then 'F' When 'JT' Then 'S' End)  and   
      LocCd in (case when OmTc<>'LCN' or OdJLWt>0 then (case when OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end) else (case when OmFrDmDc ='C' then OmToDmLoc else OmFrDmLoc end) end)
      /*LocCd in (OmFrDmLoc,OmToDmLoc)*/  
  Where   
    not Exists (select 'x' from Param Join UsrAccess on UaUsrCd = PMCd where PTyp= 'USR' and           
     PMCd= Del.ModUsr and UaMnuCd = Del.OdTc and UaSeoYn= 'Y' and PValue= OmSeoPwd)           
    and not Exists (Select 'x' from Param where PTyp= 'SEOPWD' and PMCD= 'SEOPWD' and OmSeoPwd= PValue1)               
    AND (Select CAST(SUM(DDmDrwT)-SUM(DDmCrWt) As Decimal(10,3)) from DmYyMm  
    where Del.OdPrtKEy=DPrtKey and Del.OdCoCd= DCoCd and Del.ODJLRmCd = DJLRmCd and Del.ODJLLotNo = DJLLotNo and Del.ODJLLn1 = DJLLn1 and Del.ODJLLn2 = DJLLn2 and Del.ODJLLn3 = DJLLn3 and
(case when OmTc<>'LCN' OR ODJLWT>0 THEN (CASE WHEN OmFrDmDC = 'C' then OmFrDmLoc when OmFrDmDC = 'D' then OmToDmLoc else '' end ) ELSE (CASE WHEN OMFRDMDC='D' THEN OMFRDMLOC ELSE OMTODMLOC END) END)=DLoc
     /*Del.OdDmId= DDmId and DLoc=LocCd*/  
      )< 0  
    and OmFrDmLoc <>'' and OmToDmLoc<>''  
  If @Row> 0  
  begin  
   Select @ErrMsg= 'Only Authorized User Can Enter Negative Stock For This RmCode, Lot, Sz'  
     GoTo ErrHndlr  
     end  
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

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED


