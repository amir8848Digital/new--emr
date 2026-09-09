TAKEN
----

Update Tcstr set TsLbl = 'Main Metal Wt Factor', TsStyle = 'N', TsMaxLen =11, tsmask='#0.0000000' where
tsidname = 'PNum' and tstyp='DMSZ' and tstbl='Param' 

gogo

/* insert Tcstr records for ptyp-BSW & G */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	
	
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrSTWGrp', 	
  	'N' , 	
	'BSW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'G',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='G' and TsIdName='wRrQw')



/* insert Tcstr records for ptyp-BSW & P */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrSTWGrp', 	
  	'N' , 	
	'BSW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'P',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='P' and TsIdName='wRrQw')



/* insert Tcstr records for ptyp-BSW & S */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrSTWGrp', 	
  	'N' , 	
	'BSW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'S',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='S' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-BSW & L */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrSTWGrp', 	
  	'N' , 	
	'BSW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'L',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='L' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-BSW & X */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrSTWGrp', 	
  	'N' , 	
	'BSW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BSW' ,
	'X',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BSW' and TsSTyp='X' and TsIdName='wRrQw')




/* insert Tcstr records for ptyp-STW & G */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrFixMaxTol', 	
  	'N' , 	
	'Max Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrFixMinTol', 	
  	'N' , 	
	'Min Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrSTWGrp', 	
  	'N' , 	
	'STW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'G',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='G' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-STW & P */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrFixMaxTol', 	
  	'N' , 	
	'Max Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrFixMinTol', 	
  	'N' , 	
	'Min Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrSTWGrp', 	
  	'N' , 	
	'STW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'P',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='P' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-STW & S */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrFixMaxTol', 	
  	'N' , 	
	'Max Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrFixMinTol', 	
  	'N' , 	
	'Min Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrSTWGrp', 	
  	'N' , 	
	'STW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'S',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='S' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-STW & L */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrFixMaxTol', 	
  	'N' , 	
	'Max Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrFixMinTol', 	
  	'N' , 	
	'Min Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrSTWGrp', 	
  	'N' , 	
	'STW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'L',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='L' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-STW & X */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrFixMaxTol', 	
  	'N' , 	
	'Max Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrFixMinTol', 	
  	'N' , 	
	'Min Wt' ,                   
	'F' ,
	'##0.0000',	
	8,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrSalRt', 	
  	'N' , 	
	'Weight   ' ,                   
	'F' ,
	'##0.0000',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrSTWGrp', 	
  	'N' , 	
	'STW Grp' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'#0.0000',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'STW' ,
	'X',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='STW' and TsSTyp='X' and TsIdName='wRrQw')




/* insert Tcstr records for ptyp-BS & G */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'G',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='G' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-BS & P */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'P',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='P' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-BS & S */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'S',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='S' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-BS & L */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'L',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='L' and TsIdName='wRrQw')


/* insert Tcstr records for ptyp-BS & X */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrCd')



insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrCmCurCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrCstRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrDmCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrDmCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrDmSfx')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrDmSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrFixMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrFixMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrFrLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrLabMCd')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrLsCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrOpnMaxTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'N' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrOpnMinTol')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrRmKT')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrRngSz')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrSalRt')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrSCtg')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrSTWGrp')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrToLn')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='RrXWtPer')


insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'BS' ,
	'X',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='BS' and TsSTyp='X' and TsIdName='wRrQw')



 
 