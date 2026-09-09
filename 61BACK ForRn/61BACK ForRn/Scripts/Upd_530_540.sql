SET TRANSACTION ISOLATION LEVEL READ COMMITTED

GOGO
DELETE FROM TRGDISA
GOGO

Update TcStr set TsStyle ='B' where TsTbl ='TxnRm' and TsIdName ='CmdImpExc' and TsTyp ='MV'
Update TcStr set TsStyle ='B',TsDataTyp ='U',TsMaxLen =8 where TsTbl ='TxnRm' and TsIdName ='WCpGrpOpt' and TsTyp ='PR' and TsSTyp ='X'
Update TcStr set TsLbl ='Group Option' where TsTbl ='TxnRm' and TsIdName ='LblGrpOpt' and TsTyp ='PR' and TsSTyp ='Lbl'
Update TcStr set TsLbl ='Group Option' where TsTbl ='TxnRm' and TsIdName ='LblGrpOpt' and TsTyp ='CNV' and TsSTyp ='Lbl'

GOGO

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'TxnRm'
           ,'MA'
           ,'A'
           ,''
           ,'B'
           ,''
           ,''
           ,''
           ,0
           ,'' from TcStr where not exists(select 'a' from TcStr where TsTbl ='TxnRm' and TsTyp ='MA' and TsSTyp ='A' and TsIdName ='')


INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'TxnRm'
           ,'MA'
           ,'X'
           ,'TBillDt'
           ,'B'
           ,''
           ,'D'
           ,''
           ,8
           ,'N' from TcStr where not exists(select 'a' from TcStr where TsTbl ='TxnRm' and TsTyp ='MA' and 
           TsSTyp ='X' and TsIdName ='TBillDt')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'TxnRm'
           ,'MA'
           ,'X'
           ,'TBillNo'
           ,'B'
           ,''
           ,'U'
           ,''
           ,20
           ,'N' from TcStr where not exists(select 'a' from TcStr where TsTbl ='TxnRm' and TsTyp ='MA' and 
           TsSTyp ='X' and TsIdName ='TBillNo')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'TxnRm'
           ,'MA'
           ,'X'
           ,'TBOENo'
           ,'B'
           ,''
           ,'U'
           ,''
           ,20
           ,'N' from TcStr where not exists(select 'a' from TcStr where TsTbl ='TxnRm' and TsTyp ='MA' and 
           TsSTyp ='X' and TsIdName ='TBOENo')


INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'TxnRm'
           ,'MA'
           ,'X'
           ,'TBOEDt'
           ,'B'
           ,''
           ,'D'
           ,''
           ,8
           ,'N' from TcStr where not exists(select 'a' from TcStr where TsTbl ='TxnRm' and TsTyp ='MA' and 
           TsSTyp ='X' and TsIdName ='TBOEDt')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'TxnRm'
           ,'MA'
           ,'X'
           ,'TSuppCd'
           ,'B'
           ,''
           ,'U'
           ,''
           ,8
           ,'N' from TcStr where not exists(select 'a' from TcStr where TsTbl ='TxnRm' and TsTyp ='MA' and 
           TsSTyp ='X' and TsIdName ='TSuppCd')


INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'TxnRm'
           ,'MA'
           ,'X'
           ,'WCMNAME'
           ,'B'
           ,''
           ,'U'
           ,''
           ,45
           ,'N' from TcStr where not exists(select 'a' from TcStr where TsTbl ='TxnRm' and TsTyp ='MA' and 
           TsSTyp ='X' and TsIdName ='WCMNAME')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'TxnRm'
           ,'MA'
           ,'X'
           ,'TCNVRT'
           ,'B'
           ,''
           ,'F'
           ,'##0.000'
           ,7
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='TxnRm' and TsTyp ='MA' and 
           TsSTyp ='X' and TsIdName ='TCNVRT')

GOGO


IF NOT EXISTS(SELECT * FROM   INFORMATION_SCHEMA.COLUMNS
          WHERE  upper(TABLE_NAME) = 'TXN' AND upper(COLUMN_NAME) = upper('TAck'))
BEGIN
	Alter Table Txn add TAck varchar(1) not null default ''
END

GOGO

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'CHR'
           ,''
           ,'WACKYN'
           ,'N'
           ,'AckYn'
           ,'U'
           ,''
           ,1
           ,'N' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='CHR' and TsSTyp ='' and TsIdName ='WACKYN')
GOGO


INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'CSEQ'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='CSEQ' and TsSTyp ='' and TsIdName ='WACKYN')
GOGO
INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'DEFCHR'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='DEFCHR' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'FGSUBLOC'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='FGSUBLOC' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'LINE'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='LINE' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'PRCSSEQ'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='PRCSSEQ' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'PRDSEQ'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='PRDSEQ' and TsSTyp ='' and TsIdName ='WACKYN')


INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'RWIWLONG'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='RWIWLONG' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'RWIWSEQ'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='RWIWSEQ' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'TRDEF'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='TRDEF' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'WIWLONG'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='WIWLONG' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'WIWSEQ'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='WIWSEQ' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'WORK'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='WORK' and TsSTyp ='' and TsIdName ='WACKYN')

INSERT INTO [TcStr]
           ([TsTbl]
           ,[TsTyp]
           ,[TsSTyp]
           ,[TsIdName]
           ,[TsStyle]
           ,[TsLbl]
           ,[TsDataTyp]
           ,[TsMask]
           ,[TsMaxLen]
           ,[TsBlank])
Select top 1
           'vParam'
           ,'PTYP'
           ,''
           ,'WACKYN'
           ,'I'
           ,''
           ,'U'
           ,''
           ,1
           ,'Y' from TcStr where not exists(select 'a' from TcStr where TsTbl ='vParam' and TsTyp ='PTYP' and TsSTyp ='' and TsIdName ='WACKYN')

GOGO




SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

GOGO