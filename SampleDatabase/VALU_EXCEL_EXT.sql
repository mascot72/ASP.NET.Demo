create table VALU_EXCEL_EXT(
[ID] int primary key not null,
--[ExtID] varchar(10) primary key not null,
[Name] varchar(255) not null,
[CreateDate] datetime not null default(getdate())
)
GO