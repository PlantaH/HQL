-- 1. Crear la tabla en SQL Server si no existe
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Productos]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[Productos](
        [id] [int] IDENTITY(1,1) NOT NULL,
        [nombre] [varchar](150) NOT NULL,
        [precio] [decimal](10, 2) NOT NULL,
        [stock] [int] NOT NULL DEFAULT 0,
        [activo] [bit] NOT NULL DEFAULT 1,
        CONSTRAINT [PK_Productos] PRIMARY KEY CLUSTERED ([id] ASC)
    );
END
GO

-- 2. Insertar registros de prueba
INSERT INTO [dbo].[Productos] ([nombre], [precio], [stock], [activo]) 
VALUES 
('Fideos Tallarines 500g', 1200.50, 45, 1),
('Fideos Moñito 500g', 1150.00, 30, 1),
('Salsa de Tomate Filetto 340g', 850.00, 100, 1),
('Salsa de Tomate Tuco 340g', 900.00, 20, 1),
('Queso Rallado Reggianito 100g', 1500.00, 60, 1),
('Aceite de Girasol 1L', 2100.00, 15, 1);
GO