SET TRANSACTION ISOLATION LEVEL READ COMMITTED

GOGO

update TcStr set TsStyle ='N' where TsTbl ='FgBag' and TsTyp ='Fb' and TsSTyp ='A' and TsIdName ='FdPlChr'
update TcStr set TsStyle ='N' where TsTbl ='FgBag' and TsTyp ='Fb' and TsSTyp ='A' and TsIdName ='FdPlYy'
update TcStr set TsStyle ='N' where TsTbl ='FgBag' and TsTyp ='Fb' and TsSTyp ='A' and TsIdName ='FdPlNo'

GOGO

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED


