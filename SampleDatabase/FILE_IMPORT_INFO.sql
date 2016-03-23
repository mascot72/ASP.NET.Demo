CREATE TABLE [dbo].[FILE_IMPORT_INFO] (
[ID] int PRIMARY KEY IDENTITY(1, 1) NOT NULL,
[Path] nvarchar(500) NULL,
[Name] nvarchar(500) NULL,
[ExtName] varchar(10) NULL,
[Result] varchar(255) NULL,
[Reason] nvarchar(2000) NULL,
[Remark] nvarchar(1000) NULL,
[Extend] nvarchar(1000) NULL,
[CreateDate] datetime default(getdate()) NULL,
[Creator] varchar(255) NULL,
[Size] bigint NULL)
GO