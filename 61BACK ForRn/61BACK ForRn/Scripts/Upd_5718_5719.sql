SET TRANSACTION ISOLATION LEVEL READ COMMITTED
GOGO

ALTER  View [dbo].[vwDmStkNumMaster]
As
	Select OdPrtKey, OdIdNo, OdDmCd , OdDmSz , OdSfx , OdDmStkCoCd , OdDmStkYy , OdDmStkChr , OdDmStkNo, 
				(case when OdTc <>'JRO' then 'Y' else 'N' end) as MasterYN, OrdDsg.OdOmDt 
	from OrdDsg 
				where (
		                				          OdIdNo= ( 
								  
													Select top 1 OdIdNo
																											from 
													OrdDsg ForMaster 
													join OrdMst on OmIdNo = OdOmIdNo
														where   OrdDsg.OdDmCd = ForMaster.OdDmCd and OrdDsg.OdDmSz = ForMaster.OdDmSz and OrdDsg.OdSfx = ForMaster.OdSfx and ForMaster.OdDmStkCoCd = OrdDsg.OdDmStkCoCd
																and OrdDsg.OdDmStkYy = ForMaster.OdDmStkYy and OrdDsg.OdDmStkChr = ForMaster.OdDmStkChr and OrdDsg.OdDmStkNo = ForMaster.OdDmStkNo and ForMaster.OdTc in ('JPR','JCI','JCP') and
														(Select PValue from Param where Ptyp='TC' and PMCd=OmTc)='JT'
														  order by OmDt, OdIdNo		
									      )
							    )



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
/*  select @row= count('x') from  deleted del Join OrdMst on OmIdNo=del.OdOmIdNo and OmCoCd= del.OdCoCd and OmTc= del.OdTc  
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
*/    
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
	--select @row= count('x') from inserted where UoMnuCd in ('RepJtPnd', 'RepJtPndRm') and 
	--	UoOmTc not in (Select distinct TmBaseTcs from TcMst Where TmCoCd=UoCoCdFr)
	--if @row> 0 
	--begin
	--	select @errmsg= 'Report Cannot Be Viewed For This TC'
	--	goto ErrHndlr
	--end

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
/*	select @row= count('x') from inserted ins where UoMnuCd in ('RepDsgCat') 
	 and not exists (select 'x' from param where PTyp= 'DSGEXC' and PMCd= Ins.UoDsgExc)
	if @row> 0 
	begin
		select @errmsg= 'Select if to Show All Design/Free Design /Region Exclusive/Not Exclusive'
		goto ErrHndlr
	end
*/
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
	--sv.53 OmPo group option only for invoice printing & packing list report which is mentioned in PValue	
	select @row= count('x') from inserted where UoInGrp<>'' and 
		not exists (select 'x' from Param where PTyp ='INGR' and UoInGrp= PMCd and 
						(PValue='' or '%,'+PValue +',%' like '%,'+ UoMnuCd +',%') ) 
	if @row> 0 
	begin
		select @errmsg= 'Invalid Group Combination'
		goto ErrHndlr
	end
	--sv.53 OmPo group option only for invoice printing & packing list report which is mentioned in PValue	
	select @row= count('x') from inserted where UoInGrp1<>'' and 
		not exists (select 'x' from Param where PTyp ='INGR' and UoInGrp1= PMCd and 
						(PValue='' or '%,'+PValue +',%' like '%,'+ UoMnuCd +',%') ) 
	if @row> 0 
	begin
		select @errmsg= 'Invalid Group Combination2'
		goto ErrHndlr
	end

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

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

