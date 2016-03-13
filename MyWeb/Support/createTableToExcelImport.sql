Create Table [dbo].[FILE_IMPORT_INFO] (
[ID] int PRIMARY KEY identity(1,1) not null,
[Path] nvarchar(500) NULL,
[Name] nvarchar(500) NULL,
[ExtName] varchar(3) NULL,
[Result] varchar(50) default('S') NULL,
[Reson] nvarchar(100) NULL,
[Remark] nvarchar(1000) NULL,
[Extend] nvarchar(1000) NULL,
[CreateDate] datetime default(getdate()) NULL,
[Creator] varchar(50) NULL
)
GO

Create Table VALU_EXCEL (
[ID] bigint PRIMARY KEY identity(1,1) not null,
[FileID] int not null,
[EID] varchar(50) NULL,
[InvenNo] varchar(50) NULL,
[SGNo] varchar(50) NULL,
[TID] varchar(50) NULL,
[Date] varchar(50) NULL,
[Name] nvarchar(250) NULL,
[Version] varchar(50) NULL,
[Type] varchar(50) NULL,
[DealNo] varchar(50) NULL,
[LeadNo] varchar(50) NULL,
[Comment] nvarchar(1000) NULL,
[Currency] varchar(50) NULL,
[Category] varchar(50) NULL,
[Maker] nvarchar(100) NULL,
[Model] nvarchar(100) NULL,
[Process] nvarchar(100) NULL,
[Vintage] varchar(50) NULL,
[WaferSize] int NULL,
[SerialNo] varchar(50) NULL,
[Config] varchar(50) NULL,
[Fab] varchar(50) NULL,
[Code] varchar(50) NULL,
[Location] varchar(50) NULL,
[Inspector] varchar(50) NULL,
[InspectionSummary] nvarchar(1000) NULL,
[Remark] varchar(50) NULL,
[Comment_1] nvarchar(1000) NULL,
[Period] varchar(50) NULL,
[BuyDate] varchar(50) NULL,
[SellDate] varchar(50) NULL,
[Buyer] varchar(50) NULL,
[Seller] varchar(50) NULL,
[ToolPrice_B] varchar(50) NULL,
[TotalCost_B] varchar(50) NULL,
[SGCost_B] varchar(50) NULL,
[TotalCost_S] varchar(50) NULL,
[TotalBuy] varchar(50) NULL,
[SGTotalBuy] varchar(50) NULL,
[SellPrice_E] varchar(50) NULL,
[TargetPrice] varchar(50) NULL,
[Profit] varchar(50) NULL,
[Profit_Percent] varchar(50) NULL,
[ROI] varchar(50) NULL,
[AnnualROI] varchar(50) NULL,
[DeinstallCost_B] varchar(50) NULL,
[RiggingCost_B] varchar(50) NULL,
[ShippingCost_B] varchar(50) NULL,
[PackingCost_B] varchar(50) NULL,
[InlandTruckingCost_B] varchar(50) NULL,
[Commission_B] varchar(50) NULL,
[WarehouseCost] varchar(50) NULL,
[SGInterest] varchar(50) NULL,
[InventoryAllowance] varchar(50) NULL,
[SGCommission] varchar(50) NULL,
[Task] varchar(50) NULL,
[SGOfferUSD] varchar(50) NULL,
[Ext1] nvarchar(200) NULL,
[Ext2] nvarchar(200) NULL,
[Ext3] nvarchar(200) NULL,
[Ext4] nvarchar(200) NULL,
[Ext5] nvarchar(200) NULL,
[Ext6] nvarchar(200) NULL,
[Ext7] nvarchar(200) NULL,
[Ext8] nvarchar(200) NULL,
[Ext9] nvarchar(200) NULL,
[Ext10] nvarchar(200) NULL,
[Ext11] nvarchar(200) NULL,
[Ext12] nvarchar(200) NULL,
[Ext13] nvarchar(200) NULL,
[Ext14] nvarchar(200) NULL,
[Ext15] nvarchar(200) NULL,
[Ext16] nvarchar(200) NULL,
[Ext17] nvarchar(200) NULL,
[Ext18] nvarchar(200) NULL,
[Ext19] nvarchar(200) NULL,
[Ext20] nvarchar(200) NULL,
[Ext21] nvarchar(200) NULL,
[Ext22] nvarchar(200) NULL,
[Ext23] nvarchar(200) NULL,
[Ext24] nvarchar(200) NULL,
[Ext25] nvarchar(200) NULL,
[Ext26] nvarchar(200) NULL,
[Ext27] nvarchar(200) NULL,
[Ext28] nvarchar(200) NULL,
[Ext29] nvarchar(200) NULL,
[Ext30] nvarchar(200) NULL,
[Reson] nvarchar(100) NULL,
[CreateDate] datetime default(getdate()) NOT NULL,
[Creator] varchar(50) NULL,
)
GO