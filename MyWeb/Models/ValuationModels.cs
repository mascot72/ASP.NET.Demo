using System;
using System.Data.Linq.Mapping;
using System.ComponentModel.DataAnnotations;

namespace MyWeb.Models
{
    [Table(Name = "FILE_IMPORT_INFO")]
    public class FileImport
    {
        //Primary Key column
        [Column(CanBeNull = false, DbType = "int",
        Name = "ID", IsPrimaryKey = true, IsDbGenerated = true)]
        public Int32 ID { get; set; }

        [Column(CanBeNull = true, DbType = "nvarchar(500)", Name = "Path", IsPrimaryKey = false)]
        [Display(Name = "Path")]
        public string Path { get; set; }

        [Column(CanBeNull = true, DbType = "nvarchar(500)", Name = "Name", IsPrimaryKey = false)]
        [Display(Name = "Name")]
        public string Name { get; set; }

        [Column(CanBeNull = true, DbType = "varchar(3)", Name = "ExtName", IsPrimaryKey = false)]
        [Display(Name = "ExtName")]
        public string ExtName { get; set; }

        [Column(CanBeNull = true, DbType = "varchar(50)", Name = "Result", IsPrimaryKey = false)]
        [Display(Name = "Result")]
        public string Result { get; set; }

        [Column(CanBeNull = true, DbType = "nvarchar(100)", Name = "Reason", IsPrimaryKey = false)]
        [Display(Name = "Reason")]
        public string Reason { get; set; }

        [Column(CanBeNull = true, DbType = "nvarchar(100)", Name = "Remark", IsPrimaryKey = false)]
        [Display(Name = "Remark")]
        public string Remark { get; set; }

        [Column(CanBeNull = true, DbType = "nvarchar(1000)", Name = "Extend", IsPrimaryKey = false)]
        [Display(Name = "Extend")]
        public string Extend { get; set; }

        [Column(CanBeNull = true, DbType = "datetime", Name = "CreateDate", IsPrimaryKey = false)]
        [Display(Name = "CreateDate")]
        public DateTime? CreateDate { get; set; }

        [Column(CanBeNull = true, DbType = "varchar(50)", Name = "Creator", IsPrimaryKey = false)]
        [Display(Name = "Creator")]
        public string Creator { get; set; }

    }
}

namespace MyWeb.Models.Excel
{
    [Table(Name = "VALU_EXCEL")]
    public class ValuationModels
    {
        //Primary Key column
        [Column(CanBeNull = false, DbType = "int",
        Name = "ID", IsPrimaryKey = true, IsDbGenerated = true)]
        public Int32 ID { get; set; }

        [Display(Name = "FileID")]
        [Column(CanBeNull = false, DbType = "int", IsPrimaryKey = false)]
        public int FileID { get; set; }

        //Regular columns
        [Column(CanBeNull = true, DbType = "varchar(50)", Name = "EID", IsPrimaryKey = false)]
        [Display(Name = "EID")]
        public string EId { get; set; }        

        [Display(Name = "Inven No.")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string InvenNo { get; set; }

        [Display(Name = "SG No.")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public int SGNo { get; set; }

        [Display(Name = "TID")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string TId { get; set; }

        [Display(Name = "Date")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Date { get; set; }

        [Display(Name = "Name")]
        [Column(CanBeNull = true, DbType = "nvarchar(250)", IsPrimaryKey = false)]
        public string Name { get; set; }

        [Display(Name = "Version")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Version { get; set; }

        [Display(Name = "Type")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Type { get; set; }

        [Display(Name = "Deal No")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string DealNo { get; set; }

        [Display(Name = "Lead No")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string LeadNo { get; set; }

        [Display(Name = "Comment")]
        [Column(CanBeNull = true, DbType = "nvarchar(100)", IsPrimaryKey = false)]
        public string Comment { get; set; }

        [Display(Name = "Comment_1")]
        [Column(CanBeNull = true, DbType = "nvarchar(100)", IsPrimaryKey = false)]
        public string Comment_1 { get; set; }

        [Display(Name = "Currency")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Currency { get; set; }

        [Display(Name = "Category")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Category { get; set; }

        [Display(Name = "Maker")]
        [Column(CanBeNull = true, DbType = "nvarchar(100)", IsPrimaryKey = false)]
        public string Maker { get; set; }

        [Display(Name = "Model")]
        [Column(CanBeNull = true, DbType = "nvarchar(100)", IsPrimaryKey = false)]
        public string Model { get; set; }

        [Display(Name = "Process")]
        [Column(CanBeNull = true, DbType = "nvarchar(100)", IsPrimaryKey = false)]
        public string Process { get; set; }

        [Display(Name = "Vintage")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Vintage { get; set; }

        [Display(Name = "Wafer Size")]
        [Column(CanBeNull = true, DbType = "int", IsPrimaryKey = false)]
        public int WaferSize { get; set; }

        [Display(Name = "Serial No")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string SerialNo { get; set; }

        [Display(Name = "Config")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Config { get; set; }

        [Display(Name = "Fab")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Fab { get; set; }

        [Display(Name = "Code")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Code { get; set; }

        [Display(Name = "Location")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Location { get; set; }

        [Display(Name = "Inspector")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Inspector { get; set; }

        [Display(Name = "Inspection Summary")]
        [Column(CanBeNull = true, DbType = "nvarchar(1000)", IsPrimaryKey = false)]
        public string InspectionSummary { get; set; }

        [Display(Name = "Remark")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Remark { get; set; }

        [Display(Name = "Period")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Period { get; set; }

        [Display(Name = "*Buy Date")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string BuyDate { get; set; }

        [Display(Name = "*Sell Date")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string SellDate { get; set; }

        [Display(Name = "Buyer")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Buyer { get; set; }

        [Display(Name = "Seller")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Seller { get; set; }

        [Display(Name = "Tool Price(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string ToolPriceB { get; set; }

        [Display(Name = "Total Cost(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string TotalCostB { get; set; }

        [Display(Name = "SG Cost(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string SGCostB { get; set; }

        [Display(Name = "Total Buy")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string TotalBuy { get; set; }

        [Display(Name = "SG Total Buy")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string SGTotalBuy { get; set; }

        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        [Display(Name = "*Sell Price(E)")]
        public string SellPriceE { get; set; }

        [Display(Name = "*Target Price")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string TargetPrice { get; set; }

        [Display(Name = "Profit")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Profit { get; set; }

        [Display(Name = "ROI")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string ROI { get; set; }

        [Display(Name = "Annual ROI")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string AnnualROI { get; set; }

        [Display(Name = "DeinstallCost(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string DeinstallCostB { get; set; }

        [Display(Name = "Rigging Cost(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string RiggingCostB { get; set; }

        [Display(Name = "Shipping Cost(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string ShippingCostB { get; set; }

        [Display(Name = "Packing Cost(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string PackingCostB { get; set; }

        [Display(Name = "Inland Trucking Cost(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string InlandTruckingCostB { get; set; }

        [Display(Name = "Commission(B)")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string CommissionB { get; set; }

        [Display(Name = "Warehouse Cost")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string WarehouseCost { get; set; }

        [Display(Name = "SG Interest")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string SGInterest { get; set; }

        [Display(Name = "Inventory allowance")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string InventoryAllowance { get; set; }

        [Display(Name = "SG Commission")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string SGCommission { get; set; }

        [Display(Name = "Task")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Task { get; set; }

        [Display(Name = "SG Offer USD")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string SGOfferUSD { get; set; }
        
        //Extend columns
        [Display(Name = "Ext1")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext1 { get; set; }

        [Display(Name = "Ext2")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext2 { get; set; }

        [Display(Name = "Ext3")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext3 { get; set; }

        [Display(Name = "Ext4")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext4 { get; set; }

        [Display(Name = "Ext5")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext5 { get; set; }

        [Display(Name = "Ext6")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext6 { get; set; }
        
        [Display(Name = "Ext7")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext7 { get; set; }

        [Display(Name = "Ext8")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext8 { get; set; }

        [Display(Name = "Ext9")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext9 { get; set; }

        [Display(Name = "Ext10")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext10 { get; set; }

        [Display(Name = "Ext11")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext11 { get; set; }

        [Display(Name = "Ext12")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext12 { get; set; }

        [Display(Name = "Ext13")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext13 { get; set; }

        [Display(Name = "Ext14")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext14 { get; set; }

        [Display(Name = "Ext15")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext15 { get; set; }

        [Display(Name = "Ext16")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext16 { get; set; }

        [Display(Name = "Ext17")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext17 { get; set; }

        [Display(Name = "Ext18")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext18 { get; set; }

        [Display(Name = "Ext19")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext19 { get; set; }

        [Display(Name = "Ext20")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext20 { get; set; }

        [Display(Name = "Ext21")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext21 { get; set; }
        
        [Display(Name = "Ext22")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext22 { get; set; }

        [Display(Name = "Ext23")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext23 { get; set; }

        [Display(Name = "Ext24")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext24 { get; set; }

        [Display(Name = "Ext25")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext25 { get; set; }

        [Display(Name = "Ext26")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext26 { get; set; }

        [Display(Name = "Ext27")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext27 { get; set; }
        
        [Display(Name = "Ext28")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext28 { get; set; }

        [Display(Name = "Ext29")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext29 { get; set; }

        [Display(Name = "Ext30")]
        [Column(CanBeNull = true, DbType = "nvarchar(200)", IsPrimaryKey = false)]
        public string Ext30 { get; set; }

        [Display(Name = "Reason")]
        [Column(CanBeNull = true, DbType = "varchar(100)", IsPrimaryKey = false)]
        public string Reason { get; set; }

        [Display(Name = "Creator")]
        [Column(CanBeNull = true, DbType = "varchar(50)", IsPrimaryKey = false)]
        public string Creator { get; set; }

    }
}