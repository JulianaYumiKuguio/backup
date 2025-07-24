
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tb_Transacoes](
	[Id_Transacao] [int] IDENTITY(1,1) NOT NULL,
	[Numero_Cartao] [bigint] NOT NULL,
	[Valor_Transacao] [numeric](18, 2) NOT NULL,
	[Data_Transacao] [datetime] NOT NULL,
	[Descricao] [varchar](255) NOT NULL,
	[Status] [char](1) NOT NULL,
 CONSTRAINT [PK_tb_Transacoes] PRIMARY KEY CLUSTERED 
(
	[Id_Transacao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

