



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
    -- Gera��o de Numero_Cartao (16 d�gitos, pode repetir)
    -- Multiplica por um n�mero grande para ter 16 d�gitos e garante que n�o comece com zero (ou quase zero)
    SET @NumeroCartao = CAST(RAND(CHECKSUM(NEWID())) * 8999999999999999 + 1000000000000000 AS BIGINT);

    -- Gera��o de Valor_Transacao (entre R$ 10.00 e R$ 5000.00)
    SET @ValorTransacao = ROUND(RAND(CHECKSUM(NEWID())) * (5000.00 - 10.00) + 10.00, 2);

    -- Gera��o de Data_Transacao (aleat�ria nos �ltimos 365 dias, incluindo hor�rio)
    -- Gera um n�mero de dias aleat�rio nos �ltimos 365 dias
    DECLARE @DiasAtras INT = CAST(RAND(CHECKSUM(NEWID())) * 365 AS INT);
    -- Gera horas, minutos e segundos aleat�rios
    DECLARE @Horas INT = CAST(RAND(CHECKSUM(NEWID())) * 23 AS INT);
    DECLARE @Minutos INT = CAST(RAND(CHECKSUM(NEWID())) * 59 AS INT);
    DECLARE @Segundos INT = CAST(RAND(CHECKSUM(NEWID())) * 59 AS INT);

    SET @DataTransacao = DATEADD(SECOND, @Segundos, DATEADD(MINUTE, @Minutos, DATEADD(HOUR, @Horas, DATEADD(DAY, -@DiasAtras, GETDATE()))));


    -- Gera��o de Descricao (aleat�ria de uma lista de op��es)
    DECLARE @DescricaoIndex INT = CAST(RAND(CHECKSUM(NEWID())) * 5 AS INT) + 1; -- Gera um n�mero de 1 a 5
    SET @Descricao = CASE @DescricaoIndex
                        WHEN 1 THEN 'Compra Online de Eletr�nico'
                        WHEN 2 THEN 'Pagamento de Servi�o Essencial'
                        WHEN 3 THEN 'Compra em Loja F�sica'
                        WHEN 4 THEN 'Assinatura de Software'
                        WHEN 5 THEN 'Retirada em Caixa Eletr�nico'
                        ELSE 'Transa��o Diversa' -- Fallback
                     END;

    -- Gera��o de Status (1, 2 ou 3) - N�o pode ser nulo
    DECLARE @StatusIndex INT = CAST(RAND(CHECKSUM(NEWID())) * 3 AS INT) + 1; -- Gera um n�mero de 1 a 3
    SET @Status = CAST(@StatusIndex AS CHAR(1)); -- Converte para CHAR(1)

    -- Inser��o do registro na tabela
    INSERT INTO tb_Transacoes (Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, Status)
    VALUES (@NumeroCartao, @ValorTransacao, @DataTransacao, @Descricao, @Status);

    SET @Contador = @Contador + 1;

    -- Opcional: Mostra o progresso a cada 1000 registros
    IF @Contador % 1000 = 0
    BEGIN
        PRINT 'Inseridos ' + CAST(@Contador - 1 AS VARCHAR(10)) + ' registros...';
    END
END;

PRINT 'Inser��o de 5000 registros conclu�da para tb_Transacoes.';
SET NOCOUNT OFF; -- Restaura a contagem de linhas afetadas




-- SELECT TOP 10 * FROM tb_Transacoes
-- SELECT COUNT(*) AS QTDE FROM tb_Transacoes
-- SELECT * FROM tb_Transacoes WHERE (STATUS <> 1 OR STATUS <> 2 OR STATUS <> 3)
-- SELECT NUMERO_CARTAO FROM tb_Transacoes GROUP BY NUMERO_CARTAO
-- DELETE  FROM TB_TRANSACOES