

	-- Filtrando resultados da TVF
	SELECT
		ID_Transacao,
		Data_Transacao,
		Valor_Transacao,
		CategoriaValor
	FROM
		dbo.ObterTransacoesCategorizadasPorPeriodo('2025-07-20', '2025-07-21')

	ORDER BY
		Valor_Transacao DESC;


	-- Filtrando Resultado da Function
	SELECT *, dbo.ObterCategoriaValor(valor_transacao) AS Categoria from tb_Transacoes;



	-- Filtrando View 
	SELECT * FROM [dbo].[VW_ConsultaTransacoes]
	SELECT * FROM [dbo].[VW_ConsultaTransacoesPorCategoria]


	-- Filtrando Procedure
	EXEC SP_CalcularTotalTransacoesPorPeriodo '20/07/2025 10:30:00', '28/07/2025 18:30:00','1'
