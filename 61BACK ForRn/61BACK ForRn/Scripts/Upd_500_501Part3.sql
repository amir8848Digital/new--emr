drop index nc_invfgd on invfgd

ALTER TABLE [dbo].[InvFgd] drop  CONSTRAINT [PK_InvFgd]  

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







Disable trigger param_inupdl on param
gogo
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

gogo
enable trigger param_inupdl on param




