using System;
using System.Collections.Generic;
//using System.Data.Linq.Mapping;
//using DataAnnotations = System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Linq;

namespace Excel.Domain.Entites
{
    public class Valuation
    {
        public Valuation()
        {
            //DateTime defaultDateTime = new DateTime(1900, 1, 1);
            //this.Date = defaultDateTime;
            //this.BuyDate = defaultDateTime;
            //this.SellDate = defaultDateTime;
            this.CreateDate = DateTime.Now;
            this.ExtendContent = new List<ExtendContent>();
        }
        //Primary Key column
        [Key]
        public int ID { get; set; }

        [ForeignKey("FileImport")]
        public int FileID { get; set; }

        //Regular columns
        /*
        public string EID { get; set; }

        public string InvenNo { get; set; }

        public Double? SGNo { get; set; }

        public string TID { get; set; }

        public DateTime? Date { get; set; }

        public string Name { get; set; }

        public string Version { get; set; }

        public string Type { get; set; }

        public Double? DealNo { get; set; }

        public string LeadNo { get; set; }

        public string Comment { get; set; }

        public string Comment_1 { get; set; }

        public string Currency { get; set; }

        public string Category { get; set; }

        public string Maker { get; set; }

        public string Model { get; set; }

        public string Process { get; set; }

        public string Vintage { get; set; }

        public string WaferSize { get; set; }

        public string SerialNo { get; set; }

        public string Config { get; set; }

        public string Fab { get; set; }

        public string Code { get; set; }

        public string Location { get; set; }

        public string Inspector { get; set; }

        public string InspectionSummary { get; set; }

        public string Remark { get; set; }

        public Double? Period { get; set; }

        public DateTime? BuyDate { get; set; }

        public DateTime? SellDate { get; set; }

        public string Buyer { get; set; }

        public string Seller { get; set; }

        public Double ToolPriceB { get; set; }

        public Double TotalCostB { get; set; }

        public Double SGCostB { get; set; }

        public Double TotalCostS { get; set; }

        public Double TotalBuy { get; set; }

        public Double SGTotalBuy { get; set; }

        public Double SellPriceE { get; set; }

        public Double TargetPrice { get; set; }

        public Double Profit { get; set; }

        public Double ProfitPercent { get; set; }

        public Double ROI { get; set; }

        public string AnnualROI { get; set; }

        public Double DeinstallCostB { get; set; }

        public Double RiggingCostB { get; set; }

        public Double ShippingCostB { get; set; }

        public Double PackingCostB { get; set; }

        public Double InlandTruckingCostB { get; set; }

        public Double CommissionB { get; set; }

        public Double WarehouseCost { get; set; }

        public Double SGWarehouseCost { get; set; }

        public string SGInterest { get; set; }

        public string InventoryAllowance { get; set; }

        public string SGCommission { get; set; }

        public string Task { get; set; }

        public string SGOfferUSD { get; set; }

        public Double Qty { get; set; }

        //Extend columns

        public string Ext1 { get; set; }

        public string Ext2 { get; set; }

        public string Ext3 { get; set; }

        public string Ext4 { get; set; }

        public string Ext5 { get; set; }

        public string Ext6 { get; set; }

        public string Ext7 { get; set; }

        public string Ext8 { get; set; }

        public string Ext9 { get; set; }

        public string Ext10 { get; set; }

        public string Ext11 { get; set; }

        public string Ext12 { get; set; }

        public string Ext13 { get; set; }

        public string Ext14 { get; set; }

        public string Ext15 { get; set; }

        public string Ext16 { get; set; }

        public string Ext17 { get; set; }

        public string Ext18 { get; set; }

        public string Ext19 { get; set; }

        public string Ext20 { get; set; }

        public string Ext21 { get; set; }

        public string Ext22 { get; set; }

        public string Ext23 { get; set; }

        public string Ext24 { get; set; }

        public string Ext25 { get; set; }

        public string Ext26 { get; set; }

        public string Ext27 { get; set; }

        public string Ext28 { get; set; }

        public string Ext29 { get; set; }

        public string Ext30 { get; set; }

        public string Ext31 { get; set; }

        public string Ext32 { get; set; }

        public string Ext33 { get; set; }

        public string Ext34 { get; set; }

        public string Ext35 { get; set; }

        public string Ext36 { get; set; }

        public string Ext37 { get; set; }

        public string Ext38 { get; set; }

        public string Ext39 { get; set; }

        public string Ext40 { get; set; }

        public string Ext41 { get; set; }

        public string Ext42 { get; set; }

        public string Ext43 { get; set; }

        public string Ext44 { get; set; }

        public string Ext45 { get; set; }

        public string Ext46 { get; set; }

        public string Ext47 { get; set; }

        public string Ext48 { get; set; }

        public string Ext49 { get; set; }

        public string Ext50 { get; set; }

        public int Ref1 { get; set; }

        public string Ref2 { get; set; }

        public string Reason { get; set; }

        public string Creator { get; set; }

        */

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string EID { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string InvenNo { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string SGNo { get; set; }
        
        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string TID { get; set; }

        public DateTime? Date { get; set; }

        [MaxLength(255)]
        public string Name { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Version { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Type { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double DealNo { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string PKID { get; set; }

        [MaxLength(2048), Column(TypeName = "VARCHAR")]
        public string LeadNo { get; set; }

        [MaxLength(1000), Column(TypeName = "VARCHAR")]
        public string PeriodE { get; set; }

        [MaxLength(4096)]
        public string Comment { get; set; }

        [MaxLength(4096)]
        public string Comment_1 { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Currency { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Category { get; set; }

        [MaxLength(255)]
        public string Maker { get; set; }

        [MaxLength(255)]
        public string Model { get; set; }

        [MaxLength(255)]
        public string Process { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Vintage { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string WaferSize { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string SerialNo { get; set; }

        [MaxLength(2048), Column(TypeName = "VARCHAR")]
        public string Config { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Fab { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Code { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Location { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Inspector { get; set; }

        [MaxLength(1000)]
        public string InspectionSummary { get; set; }

        [MaxLength(1000)]
        public string Remark { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double Period { get; set; }
        
        public DateTime? BuyDate { get; set; }
        
        public DateTime? SellDate { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Buyer { get; set; }
        
        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Seller { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double ToolPriceB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double TotalCostB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double SGCostB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double TotalCostS { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double TotalBuy { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double SGTotalBuy { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double SellPriceE { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double TargetPrice { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double Profit { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double ProfitPercent { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double ROI { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string AnnualROI { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double DeinstallCostB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double RiggingCostB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double ShippingCostB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double PackingCostB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double InlandTruckingCostB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double CommissionB { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double WarehouseCost { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double SGWarehouseCost { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string SGInterest { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string InventoryAllowance { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string SGCommission { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Task { get; set; }

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string SGOfferUSD { get; set; }

        [Column(TypeName = "FLOAT")]
        public Double Qty { get; set; }

        //Extend columns
        [MaxLength(255)]
        public string Ext1 { get; set; }

        [MaxLength(255)]
        public string Ext2 { get; set; }

        [MaxLength(255)]
        public string Ext3 { get; set; }

        [MaxLength(255)]
        public string Ext4 { get; set; }

        [MaxLength(255)]
        public string Ext5 { get; set; }

        [MaxLength(255)]
        public string Ext6 { get; set; }

        [MaxLength(255)]
        public string Ext7 { get; set; }

        [MaxLength(255)]
        public string Ext8 { get; set; }

        [MaxLength(255)]
        public string Ext9 { get; set; }

        [MaxLength(255)]
        public string Ext10 { get; set; }

        [MaxLength(255)]
        public string Ext11 { get; set; }

        [MaxLength(255)]
        public string Ext12 { get; set; }

        [MaxLength(255)]
        public string Ext13 { get; set; }

        [MaxLength(255)]
        public string Ext14 { get; set; }

        [MaxLength(255)]
        public string Ext15 { get; set; }

        [MaxLength(255)]
        public string Ext16 { get; set; }

        [MaxLength(255)]
        public string Ext17 { get; set; }

        [MaxLength(255)]
        public string Ext18 { get; set; }

        [MaxLength(255)]
        public string Ext19 { get; set; }

        [MaxLength(255)]
        public string Ext20 { get; set; }

        [MaxLength(255)]
        public string Ext21 { get; set; }

        [MaxLength(255)]
        public string Ext22 { get; set; }

        [MaxLength(255)]
        public string Ext23 { get; set; }

        [MaxLength(255)]
        public string Ext24 { get; set; }

        [MaxLength(255)]
        public string Ext25 { get; set; }

        [MaxLength(255)]
        public string Ext26 { get; set; }

        [MaxLength(255)]
        public string Ext27 { get; set; }

        [MaxLength(255)]
        public string Ext28 { get; set; }

        [MaxLength(255)]
        public string Ext29 { get; set; }

        [MaxLength(255)]
        public string Ext30 { get; set; }

        [MaxLength(255)]
        public string Ext31 { get; set; }

        [MaxLength(255)]
        public string Ext32 { get; set; }

        [MaxLength(255)]
        public string Ext33 { get; set; }

        [MaxLength(255)]
        public string Ext34 { get; set; }

        [MaxLength(255)]
        public string Ext35 { get; set; }

        [MaxLength(255)]
        public string Ext36 { get; set; }

        [MaxLength(255)]
        public string Ext37 { get; set; }

        [MaxLength(255)]
        public string Ext38 { get; set; }

        [MaxLength(255)]
        public string Ext39 { get; set; }

        [MaxLength(255)]
        public string Ext40 { get; set; }

        [MaxLength(255)]
        public string Ext41 { get; set; }

        [MaxLength(255)]
        public string Ext42 { get; set; }

        [MaxLength(255)]
        public string Ext43 { get; set; }

        [MaxLength(255)]
        public string Ext44 { get; set; }

        [MaxLength(255)]
        public string Ext45 { get; set; }

        [MaxLength(255)]
        public string Ext46 { get; set; }

        [MaxLength(255)]
        public string Ext47 { get; set; }

        [MaxLength(255)]
        public string Ext48 { get; set; }

        [MaxLength(255)]
        public string Ext49 { get; set; }

        [MaxLength(255)]
        public string Ext50 { get; set; }

        public int Ref1 { get; set; }

        [MaxLength(200), Column(TypeName = "VARCHAR")]
        public string Ref2 { get; set; }

        [MaxLength(2048), Column(TypeName = "NVARCHAR")]
        public string Reason { get; set; }

        public DateTime? CreateDate { get; set; }        

        [MaxLength(255), Column(TypeName = "VARCHAR")]
        public string Creator { get; set; }

        public virtual FileImport FileImport { get; set; }

        public ICollection<ExtendContent> ExtendContent { get; set; }

    }
}