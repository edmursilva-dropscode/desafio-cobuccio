USE [master]
GO
/****** Object:  Database [AdministradoraCC]    Script Date: 19/09/2024 20:55:29 ******/
CREATE DATABASE [AdministradoraCC]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'AdministradoraCC', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\AdministradoraCC.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'AdministradoraCC_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\AdministradoraCC_log.ldf' , SIZE = 8192KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT, LEDGER = OFF
GO
ALTER DATABASE [AdministradoraCC] SET COMPATIBILITY_LEVEL = 160
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [AdministradoraCC].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [AdministradoraCC] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [AdministradoraCC] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [AdministradoraCC] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [AdministradoraCC] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [AdministradoraCC] SET ARITHABORT OFF 
GO
ALTER DATABASE [AdministradoraCC] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [AdministradoraCC] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [AdministradoraCC] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [AdministradoraCC] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [AdministradoraCC] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [AdministradoraCC] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [AdministradoraCC] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [AdministradoraCC] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [AdministradoraCC] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [AdministradoraCC] SET  DISABLE_BROKER 
GO
ALTER DATABASE [AdministradoraCC] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [AdministradoraCC] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [AdministradoraCC] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [AdministradoraCC] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [AdministradoraCC] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [AdministradoraCC] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [AdministradoraCC] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [AdministradoraCC] SET RECOVERY FULL 
GO
ALTER DATABASE [AdministradoraCC] SET  MULTI_USER 
GO
ALTER DATABASE [AdministradoraCC] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [AdministradoraCC] SET DB_CHAINING OFF 
GO
ALTER DATABASE [AdministradoraCC] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [AdministradoraCC] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [AdministradoraCC] SET DELAYED_DURABILITY = DISABLED 
GO
ALTER DATABASE [AdministradoraCC] SET ACCELERATED_DATABASE_RECOVERY = OFF  
GO
EXEC sys.sp_db_vardecimal_storage_format N'AdministradoraCC', N'ON'
GO
ALTER DATABASE [AdministradoraCC] SET QUERY_STORE = ON
GO
ALTER DATABASE [AdministradoraCC] SET QUERY_STORE (OPERATION_MODE = READ_WRITE, CLEANUP_POLICY = (STALE_QUERY_THRESHOLD_DAYS = 30), DATA_FLUSH_INTERVAL_SECONDS = 900, INTERVAL_LENGTH_MINUTES = 60, MAX_STORAGE_SIZE_MB = 1000, QUERY_CAPTURE_MODE = AUTO, SIZE_BASED_CLEANUP_MODE = AUTO, MAX_PLANS_PER_QUERY = 200, WAIT_STATS_CAPTURE_MODE = ON)
GO
USE [AdministradoraCC]
GO
/****** Object:  UserDefinedFunction [dbo].[CategoriaTransacao]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[CategoriaTransacao] (@Valor DECIMAL(10,2))
RETURNS VARCHAR(10)
AS
BEGIN
    DECLARE @Categoria VARCHAR(10)
    IF @Valor > 1000
        SET @Categoria = 'Alta'
    ELSE IF @Valor >= 500
        SET @Categoria = 'Média'
    ELSE
        SET @Categoria = 'Baixa'
    RETURN @Categoria
END;
GO
/****** Object:  Table [dbo].[Clientes]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Clientes](
	[ID_Cliente] [int] IDENTITY(1,1) NOT NULL,
	[Nome_Cliente] [varchar](100) NOT NULL,
	[Numero_Cartao] [varchar](16) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[ID_Cliente] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Transacoes]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Transacoes](
	[ID_Transacao] [int] IDENTITY(1,1) NOT NULL,
	[Numero_Cartao] [varchar](16) NOT NULL,
	[Valor_Transacao] [decimal](10, 2) NOT NULL,
	[Data_Transacao] [datetime] NOT NULL,
	[Descricao] [varchar](255) NULL,
PRIMARY KEY CLUSTERED 
(
	[ID_Transacao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  View [dbo].[vw_Transacoes]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE VIEW [dbo].[vw_Transacoes] AS
SELECT 
	t.ID_Transacao,
	c.ID_Cliente,
    c.Nome_Cliente,
    t.Numero_Cartao,
	FORMAT(t.Valor_Transacao, 'C', 'pt-BR') AS Valor_Transacao,
    t.Data_Transacao,
    t.Descricao,
    dbo.CategoriaTransacao(t.Valor_Transacao) AS Categoria
FROM 
    Transacoes t
    INNER JOIN Clientes c ON t.Numero_Cartao = c.Numero_Cartao;
GO
/****** Object:  View [dbo].[vw_Estatisticas]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO





CREATE VIEW [dbo].[vw_Estatisticas] AS
SELECT
    c.Numero_Cartao,
    c.Nome_Cliente,
    SUM(t.Valor_Transacao) AS Valor_Total,
    COUNT(t.ID_Transacao) AS Quantidade_Transacoes
FROM
    Transacoes t
INNER JOIN
    Clientes c ON t.Numero_Cartao = c.Numero_Cartao
GROUP BY
    c.Numero_Cartao,
    c.Nome_Cliente;
GO
SET IDENTITY_INSERT [dbo].[Clientes] ON 

INSERT [dbo].[Clientes] ([ID_Cliente], [Nome_Cliente], [Numero_Cartao]) VALUES (1, N'João Silva', N'1234567890123456')
INSERT [dbo].[Clientes] ([ID_Cliente], [Nome_Cliente], [Numero_Cartao]) VALUES (2, N'Maria Santos', N'9876543210987654')
INSERT [dbo].[Clientes] ([ID_Cliente], [Nome_Cliente], [Numero_Cartao]) VALUES (3, N'Guilhrme gomes', N'1212121213134545')
INSERT [dbo].[Clientes] ([ID_Cliente], [Nome_Cliente], [Numero_Cartao]) VALUES (4, N'Roberto Flix', N'3434343456566778')
INSERT [dbo].[Clientes] ([ID_Cliente], [Nome_Cliente], [Numero_Cartao]) VALUES (5, N'Roberta Baldim Galassi', N'1245567898324565')
SET IDENTITY_INSERT [dbo].[Clientes] OFF
GO
SET IDENTITY_INSERT [dbo].[Transacoes] ON 

INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_Cartao], [Valor_Transacao], [Data_Transacao], [Descricao]) VALUES (10, N'1212121213134545', CAST(1500.35 AS Decimal(10, 2)), CAST(N'2024-09-18T00:00:00.000' AS DateTime), N'Pagamento Mecânica Marção LTDA')
INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_Cartao], [Valor_Transacao], [Data_Transacao], [Descricao]) VALUES (11, N'1234567890123456', CAST(550.00 AS Decimal(10, 2)), CAST(N'2024-09-17T00:00:00.000' AS DateTime), N'PAgamento do material de contrução empresa TERMINAL LTDA')
INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_Cartao], [Valor_Transacao], [Data_Transacao], [Descricao]) VALUES (12, N'1245567898324565', CAST(2500.00 AS Decimal(10, 2)), CAST(N'2024-09-16T00:00:00.000' AS DateTime), N'Pagamento album de formaruda empresa FONTESLUZ LTDA')
INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_Cartao], [Valor_Transacao], [Data_Transacao], [Descricao]) VALUES (13, N'3434343456566778', CAST(4458.33 AS Decimal(10, 2)), CAST(N'2024-09-19T00:00:00.000' AS DateTime), N'Pagamento fornecedor LATIM LTDA')
INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_Cartao], [Valor_Transacao], [Data_Transacao], [Descricao]) VALUES (14, N'9876543210987654', CAST(545.00 AS Decimal(10, 2)), CAST(N'2024-09-18T00:00:00.000' AS DateTime), N'Pagamento fornecedor BLINK LTDA')
INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_Cartao], [Valor_Transacao], [Data_Transacao], [Descricao]) VALUES (15, N'1212121213134545', CAST(100.00 AS Decimal(10, 2)), CAST(N'2024-09-18T00:00:00.000' AS DateTime), N'Paramento fornecedor JOVEM LTDA')
SET IDENTITY_INSERT [dbo].[Transacoes] OFF
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [UQ__Clientes__A28F09A9721EE814]    Script Date: 19/09/2024 20:55:29 ******/
ALTER TABLE [dbo].[Clientes] ADD UNIQUE NONCLUSTERED 
(
	[Numero_Cartao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Transacoes]  WITH CHECK ADD FOREIGN KEY([Numero_Cartao])
REFERENCES [dbo].[Clientes] ([Numero_Cartao])
GO
/****** Object:  StoredProcedure [dbo].[sp_AtualizarCliente]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_AtualizarCliente]
    @ID_Cliente INT,
    @Nome_Cliente VARCHAR(100),
    @Numero_Cartao VARCHAR(16)
AS
BEGIN
    UPDATE Clientes
    SET Nome_Cliente = @Nome_Cliente, Numero_Cartao = @Numero_Cartao
    WHERE ID_Cliente = @ID_Cliente;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_AtualizarTransacao]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_AtualizarTransacao]
    @ID_Transacao INT,
    @Numero_Cartao VARCHAR(16),
    @Data_Transacao DATETIME,
    @Valor_Transacao DECIMAL(10, 2),
    @Descricao VARCHAR(255)
AS
BEGIN
    UPDATE Transacoes
    SET Numero_Cartao = @Numero_Cartao, Data_Transacao = @Data_Transacao, Valor_Transacao = @Valor_Transacao, Descricao = @Descricao
    WHERE ID_Transacao = @ID_Transacao;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ConsultarTransacoes]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_ConsultarTransacoes]
    @Numero_Cartao VARCHAR(16) = NULL,
    @Data_Transacao DATETIME = NULL,
    @Valor_Transacao DECIMAL(10, 2) = NULL
AS
BEGIN
    SELECT ID_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao
    FROM Transacoes
    WHERE (@Numero_Cartao IS NULL OR Numero_Cartao = @Numero_Cartao)
      AND (@Data_Transacao IS NULL OR Data_Transacao = @Data_Transacao)
      AND (@Valor_Transacao IS NULL OR Valor_Transacao = @Valor_Transacao);
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ExcluirCliente]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_ExcluirCliente]
    @ID_Cliente INT
AS
BEGIN
    DELETE FROM Clientes WHERE ID_Cliente = @ID_Cliente;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ExcluirTransacao]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_ExcluirTransacao]
    @ID_Transacao INT
AS
BEGIN
    DELETE FROM Transacoes WHERE ID_Transacao = @ID_Transacao;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_InserirCliente]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_InserirCliente]
    @Nome_Cliente VARCHAR(100),
    @Numero_Cartao VARCHAR(16)
AS
BEGIN
    INSERT INTO Clientes (Nome_Cliente, Numero_Cartao)
    VALUES (@Nome_Cliente, @Numero_Cartao);
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_InserirTransacao]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[sp_InserirTransacao]
    @Numero_Cartao VARCHAR(16),
    @Data_Transacao DATETIME,
    @Valor_Transacao DECIMAL(10, 2),
    @Descricao VARCHAR(255)
AS
BEGIN
    INSERT INTO Transacoes (Numero_Cartao, Data_Transacao, Valor_Transacao, Descricao)
    VALUES (@Numero_Cartao, @Data_Transacao, @Valor_Transacao, @Descricao);
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ListarClientes]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[sp_ListarClientes]
    @cmbLocalizar INT
AS
BEGIN
    SELECT ID_Cliente, 
           Nome_Cliente, 
           Numero_Cartao
    FROM Clientes (NOLOCK)
    ORDER BY 
        CASE 
            WHEN @cmbLocalizar = 0 THEN CAST(ID_Cliente AS VARCHAR) 
            WHEN @cmbLocalizar = 1 THEN Nome_Cliente
        END
END

GO
/****** Object:  StoredProcedure [dbo].[sp_TotalTransacoesPorPeriodo]    Script Date: 19/09/2024 20:55:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Criar stored procedure para calcular total de transações por período
CREATE PROCEDURE [dbo].[sp_TotalTransacoesPorPeriodo]
    @Data_Inicial DATE,
    @Data_Final DATE
AS
BEGIN
    SELECT 
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,
        COUNT(*) AS Quantidade_Transacoes
    FROM 
        Transacoes
    WHERE 
        Data_Transacao BETWEEN @Data_Inicial AND @Data_Final
    GROUP BY 
        Numero_Cartao
END;
GO
USE [master]
GO
ALTER DATABASE [AdministradoraCC] SET  READ_WRITE 
GO
