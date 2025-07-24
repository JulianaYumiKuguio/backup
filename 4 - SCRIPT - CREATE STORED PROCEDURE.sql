SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[SP_CalcularTotalTransacoesPorPeriodo]
    @Data_Inicial DATETIME,
    @Data_Final DATETIME,
    @Status_Transacao VARCHAR(50) = NULL -- NULL permite buscar todos os status, se n�o especificado
AS
BEGIN
    SET NOCOUNT ON; -- Evita que o n�mero de linhas afetadas seja retornado, otimizando o desempenho

    SELECT
        t.Numero_Cartao,
        SUM(t.Valor_Transacao) AS Valor_Total,
        COUNT(t.Id_Transacao) AS Quantidade_Transacoes,
        t.Status AS Status_Transacao
    FROM
        dbo.tb_Transacoes AS t
    WHERE
        t.Data_Transacao >=  CONVERT(DATETIME,@Data_Inicial, 103)
        AND t.Data_Transacao < DATEADD(DAY, 1,  CONVERT(DATETIME,@Data_Final, 103)) -- Garante que inclui todas as transa��es do Data_Final
        AND (@Status_Transacao IS NULL OR t.Status = @Status_Transacao)
    GROUP BY
        t.Numero_Cartao,
        t.Status
    ORDER BY
        t.Numero_Cartao,
        t.Status;

END

GO


