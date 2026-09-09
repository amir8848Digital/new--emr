/* insert Tcstr records for ptyp-RSZ & G */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrRngSz', 	
  	'N' , 	
	'Range Size' ,                   
	'U' ,
	'',	
	12 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'G',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='G' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-RSZ & P */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrRngSz', 	
  	'N' , 	
	'Range Size' ,                   
	'U' ,
	'',	
	12 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'P',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='P' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-RSZ & S */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrRngSz', 	
  	'N' , 	
	'Range Size' ,                   
	'U' ,
	'',	
	12 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'S',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='S' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-RSZ & L */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrRngSz', 	
  	'N' , 	
	'Range Size' ,                   
	'U' ,
	'',	
	12 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'L',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='L' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-RSZ & X */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrRngSz', 	
  	'N' , 	
	'Range Size' ,                   
	'U' ,
	'',	
	12 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'X',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='X' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-RSZ & M */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrFrLn', 	
  	'N' , 	
	'Fr Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrRngSz', 	
  	'N' , 	
	'Range Size' ,                   
	'U' ,
	'',	
	12 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrToLn', 	
  	'N' , 	
	'To Size' ,                   
	'F' ,
	'##0.0000',	
	8 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'RSZ' ,
	'M',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='RSZ' and TsSTyp='M' and TsIdName='wRrQw')
gogo




/* insert Tcstr records for ptyp-INC & G */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrFrLn', 	
  	'I' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrToLn', 	
  	'I' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'G',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='G' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INC & P */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrFrLn', 	
  	'I' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrToLn', 	
  	'I' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'P',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='P' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INC & S */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrFrLn', 	
  	'I' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrToLn', 	
  	'I' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'S',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='S' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INC & L */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrFrLn', 	
  	'I' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrToLn', 	
  	'I' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'L',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='L' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INC & X */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrFrLn', 	
  	'I' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrToLn', 	
  	'I' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'X',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='X' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INC & M */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrFrLn', 	
  	'I' , 	
	'Fr Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrToLn', 	
  	'I' , 	
	'To Size' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INC' ,
	'M',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INC' and TsSTyp='M' and TsIdName='wRrQw')
gogo



/* insert Tcstr records for ptyp-INP & D */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'D',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='D' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'D',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='D' and TsIdName='RrRmKT')
gogo

/* insert Tcstr records for ptyp-INP & C */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'C',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='C' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'C',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='C' and TsIdName='RrRmKT')
gogo


/* insert Tcstr records for ptyp-INP & G */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'G',
	'wRrQw', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='G' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INP & P */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'P',
	'wRrQw', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='P' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INP & S */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'S',
	'wRrQw', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='S' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INP & L */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'L',
	'wRrQw', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='L' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INP & X */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'X',
	'wRrQw', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='X' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-INP & M */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrSalRt', 	
  	'N' , 	
	'Inc %' ,                   
	'F' ,
	'-##0.00',	
	7 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'INP' ,
	'M',
	'wRrQw', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='INP' and TsSTyp='M' and TsIdName='wRrQw')
gogo




/* insert Tcstr records for ptyp-PTR & G */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'G',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='G' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-PTR & P */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'P',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='P' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-PTR & S */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'S',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='S' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-PTR & L */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'L',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='L' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-PTR & X */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'X',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='X' and TsIdName='wRrQw')
gogo

/* insert Tcstr records for ptyp-PTR & M */
insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	16 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrCmCurCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrCmCurCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrCstRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrCstRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrDmCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	15 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrDmCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrDmCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrDmCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrDmSfx', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrDmSfx')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrDmSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrDmSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrFixMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrFixMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrFixMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrFixMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrFrLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrFrLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrLabMCd', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrLabMCd')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrLsCtg', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrLsCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrOpnMaxTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrOpnMaxTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrOpnMinTol', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0',	
	3 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrOpnMinTol')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrRmKT', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrRmKT')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrRngSz', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	10 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrRngSz')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrSalRt', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'######0.000',	
	11 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrSalRt')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrSCtg', 	
  	'N' , 	
	'Sub Ctg' ,                   
	'U' ,
	'',	
	5 ,                                
	'N' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrSCtg')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrSTWGrp', 	
  	'I' , 	
	'' ,                   
	'U' ,
	'',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrSTWGrp')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrToLn', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'#0.00',	
	5 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrToLn')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'RrXWtPer', 	
  	'I' , 	
	'' ,                   
	'F' ,
	'##0.00',	
	6 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='RrXWtPer')
gogo

insert into TcStr 
	(TsTbl,
	TsTyp   ,  
	TsSTyp    , 
	TsIdName     ,
	TsStyle    ,                     
	TsLbl ,                
    TsDataTyp   ,                     
	TsMask     ,             
	TsMaxLen  ,                                                                                         
	TsBlank  )
select top 1 
	'RmRt' ,  	
	'PTR' ,
	'M',
	'wRrQw', 	
  	'I' , 	
	'Q/W Sal' ,                   
	'U' ,
	'',	
	1 ,                                
	'Y' 
 from TcStr 
 where not exists (Select 'a' from TcStr where TsTbl='RmRt' and TsTyp='PTR' and TsSTyp='M' and TsIdName='wRrQw')
gogo

