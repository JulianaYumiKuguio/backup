SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[ObterCategoriaValor]
(
    @Valor DECIMAL(18, 2) -- O valor de entrada (use DECIMAL para flexibilidade)
)
RETURNS VARCHAR(20) -- O tipo de dado do retorno (a categoria em texto)
AS
BEGIN
    DECLARE @Categoria VARCHAR(20);

    SET @Categoria = CASE
                        WHEN @Valor > 2000 THEN 'Premium'
                        WHEN @Valor >= 1000 AND @Valor <= 2000 THEN 'Alta'
                        WHEN @Valor >= 500 AND @Valor < 1000 THEN 'Média'
                        WHEN @Valor < 500 THEN 'Baixa'
                        ELSE 'Não Classificado' -- Para valores NULL ou outros casos
                     END;

    RETURN @Categoria;
END;
GO


