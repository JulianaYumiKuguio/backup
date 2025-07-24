



SET NOCOUNT ON; -- Desabilita a contagem de linhas afetadas para cada INSERT, acelerando o processo

DECLARE @Contador INT = 1;
DECLARE @NumeroCartao BIGINT;
DECLARE @ValorTransacao NUMERIC(18, 2);
DECLARE @DataTransacao DATETIME;
DECLARE @Descricao VARCHAR(255);
DECLARE @Status CHAR(1);

-- Loop para inserir 5.000 registros
WHILE @Contador <= 5000
BEGIN
    -- Geração de Numero_Cartao (16 dígitos, pode repetir)
    -- Multiplica por um número grande para ter 16 dígitos e garante que não comece com zero (ou quase zero)
    SET @NumeroCartao = CAST(RAND(CHECKSUM(NEWID())) * 8999999999999999 + 1000000000000000 AS BIGINT);

    -- Geração de Valor_Transacao (entre R$ 10.00 e R$ 5000.00)
    SET @ValorTransacao = ROUND(RAND(CHECKSUM(NEWID())) * (5000.00 - 10.00) + 10.00, 2);

    -- Geração de Data_Transacao (aleatória nos últimos 365 dias, incluindo horário)
    -- Gera um número de dias aleatório nos últimos 365 dias
    DECLARE @DiasAtras INT = CAST(RAND(CHECKSUM(NEWID())) * 365 AS INT);
    -- Gera horas, minutos e segundos aleatórios
    DECLARE @Horas INT = CAST(RAND(CHECKSUM(NEWID())) * 23 AS INT);
    DECLARE @Minutos INT = CAST(RAND(CHECKSUM(NEWID())) * 59 AS INT);
    DECLARE @Segundos INT = CAST(RAND(CHECKSUM(NEWID())) * 59 AS INT);

    SET @DataTransacao = DATEADD(SECOND, @Segundos, DATEADD(MINUTE, @Minutos, DATEADD(HOUR, @Horas, DATEADD(DAY, -@DiasAtras, GETDATE()))));


    -- Geração de Descricao (aleatória de uma lista de opções)
    DECLARE @DescricaoIndex INT = CAST(RAND(CHECKSUM(NEWID())) * 5 AS INT) + 1; -- Gera um número de 1 a 5
    SET @Descricao = CASE @DescricaoIndex
                        WHEN 1 THEN 'Compra Online de Eletrônico'
                        WHEN 2 THEN 'Pagamento de Serviço Essencial'
                        WHEN 3 THEN 'Compra em Loja Física'
                        WHEN 4 THEN 'Assinatura de Software'
                        WHEN 5 THEN 'Retirada em Caixa Eletrônico'
                        ELSE 'Transação Diversa' -- Fallback
                     END;

    -- Geração de Status (1, 2 ou 3) - Não pode ser nulo
    DECLARE @StatusIndex INT = CAST(RAND(CHECKSUM(NEWID())) * 3 AS INT) + 1; -- Gera um número de 1 a 3
    SET @Status = CAST(@StatusIndex AS CHAR(1)); -- Converte para CHAR(1)

    -- Inserção do registro na tabela
    INSERT INTO tb_Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status)
    VALUES (@NumeroCartao, @ValorTransacao, @DataTransacao, @Descricao, @Status);

    SET @Contador = @Contador + 1;

    -- Opcional: Mostra o progresso a cada 1000 registros
    IF @Contador % 1000 = 0
    BEGIN
        PRINT 'Inseridos ' + CAST(@Contador - 1 AS VARCHAR(10)) + ' registros...';
    END
END;

PRINT 'Inserção de 5000 registros concluída para tb_Transacoes.';
SET NOCOUNT OFF; -- Restaura a contagem de linhas afetadas




-- SELECT TOP 10 * FROM tb_Transacoes
-- SELECT COUNT(*) AS QTDE FROM tb_Transacoes
-- SELECT * FROM tb_Transacoes WHERE (STATUS <> 1 OR STATUS <> 2 OR STATUS <> 3)
-- SELECT NUMERO_CARTAO FROM tb_Transacoes GROUP BY NUMERO_CARTAO
-- DELETE  FROM TB_TRANSACOES