﻿CREATE TABLE [dbo].[VALU_EXCEL] (
[ID] int PRIMARY KEY IDENTITY(1, 1) NOT NULL,
[FileID] int NOT NULL,
[EID] varchar(255) NULL,
[InvenNo] varchar(255) NULL,
[SGNo] float NULL,
[TID] varchar(255) NULL,
[Date] datetime NULL,
[Name] nvarchar(255) NULL,
[Version] varchar(255) NULL,
[Type] varchar(255) NULL,
[DealNo] float NULL,
[LeadNo] varchar(255) NULL,
[Comment] nvarchar(2000) NULL,
[Currency] varchar(255) NULL,
[Category] varchar(255) NULL,
[Maker] nvarchar(100) NULL,
[Model] nvarchar(100) NULL,
[Process] nvarchar(100) NULL,
[Vintage] varchar(255) NULL,
[WaferSize] varchar(255) NULL,
[SerialNo] varchar(255) NULL,
[Config] varchar(255) NULL,
[Fab] varchar(255) NULL,
[Code] varchar(255) NULL,
[Location] varchar(255) NULL,
[Inspector] varchar(255) NULL,
[InspectionSummary] nvarchar(1000) NULL,
[Remark] nvarchar(1000) NULL,
[Comment_1] nvarchar(1000) NULL,
[Period] float NULL,
[BuyDate] datetime NULL,
[SellDate] datetime NULL,
[Buyer] varchar(255) NULL,
[Seller] varchar(255) NULL,
[ToolPriceB] float NULL,
[TotalCostB] float NULL,
[SGCostB] float NULL,
[TotalCostS] float NULL,
[TotalBuy] float NULL,
[SGTotalBuy] float NULL,
[SellPriceE] float NULL,
[TargetPrice] float NULL,
[Profit] float NULL,
[ProfitPercent] float NULL,
[ROI] float NULL,
[AnnualROI] varchar(255) NULL,
[DeinstallCostB] float NULL,
[RiggingCostB] float NULL,
[ShippingCostB] float NULL,
[PackingCostB] float NULL,
[InlandTruckingCostB] float NULL,
[CommissionB] float NULL,
[WarehouseCost] float NULL,
[SGWarehouseCost] float NULL,
[SGInterest] varchar(255) NULL,
[InventoryAllowance] varchar(255) NULL,
[SGCommission] varchar(255) NULL,
[Task] varchar(255) NULL,
[SGOfferUSD] varchar(255) NULL,
[Qty] float NULL,
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
[Ext31] nvarchar(200) NULL,
[Ext32] nvarchar(200) NULL,
[Ext33] nvarchar(200) NULL,
[Ext34] nvarchar(200) NULL,
[Ext35] nvarchar(200) NULL,
[Ext36] nvarchar(200) NULL,
[Ext37] nvarchar(200) NULL,
[Ext38] nvarchar(200) NULL,
[Ext39] nvarchar(200) NULL,
[Ext40] nvarchar(200) NULL,
[Ext41] nvarchar(200) NULL,
[Ext42] nvarchar(200) NULL,
[Ext43] nvarchar(200) NULL,
[Ext44] nvarchar(200) NULL,
[Ext45] nvarchar(200) NULL,
[Ext46] nvarchar(200) NULL,
[Ext47] nvarchar(200) NULL,
[Ext48] nvarchar(200) NULL,
[Ext49] nvarchar(200) NULL,
[Ext50] nvarchar(200) NULL,
[Ref1] [int] NULL,
[Ref2] [varchar](200) NULL,
[Reason] nvarchar(2000) NULL,
[CreateDate] datetime NOT NULL default(getdate()),
[Creator] varchar(255) NULL, 
    CONSTRAINT [FK_VALU_EXCEL_FILE_IMPORT_INFO] FOREIGN KEY ([FileID]) REFERENCES [FILE_IMPORT_INFO]([ID]))