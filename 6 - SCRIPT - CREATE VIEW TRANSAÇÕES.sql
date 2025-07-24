SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[VW_ConsultaTransacoes] AS


	select 
	Id_Transacao as ID, 
	Numero_Cartao, 
	Valor_Transacao, 
	Data_Transacao, 
	Descricao, 
	CASE
		WHEN Status = 1 THEN 'Aprovada' 
		WHEN Status = 2 THEN 'Pendente' 
		ELSE 'Cancelada' END
	as Status
	from [dbo].[tb_Transacoes]



GO


