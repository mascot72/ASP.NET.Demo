using System;
using System.Data.Linq.Mapping;
using System.ComponentModel.DataAnnotations;

namespace MyWeb.Models
{
    [Table(Name = "VALU_EXCEL")]
    public class ValuationModels
    {
        //Primary Key column
        [Column(CanBeNull = false, DbType = "int",
        Name = "ID", IsPrimaryKey = true)]
        public Int32 ID { get; set; }

        //Regular columns
        [Display(Name = "EID")]
        public string EId { get; set; }
        [Display(Name = "Inven No.")]
        public string InvenNo { get; set; }
        [Display(Name = "SG No.")]
        public int SGNo { get; set; }
        [Display(Name = "TID")]
        public string TId { get; set; }
        [Display(Name = "Date")]
        public string Date { get; set; }
        [Display(Name = "Name")]
        public string Name { get; set; }
        [Display(Name = "Version")]
        public string Version { get; set; }
        [Display(Name = "Type")]
        public string Type { get; set; }
        [Display(Name = "Deal No")]
        public string DealNo { get; set; }
        [Display(Name = "Lead No")]
        public string LeadNo { get; set; }
        [Display(Name = "Comment")]
        public string Comment { get; set; }
        [Display(Name = "Currency")]
        public string Currency { get; set; }
        [Display(Name = "Category")]
        public string Category { get; set; }
        [Display(Name = "Maker")]
        public string Maker { get; set; }
        [Display(Name = "Model")]
        public string Model { get; set; }
        [Display(Name = "Process")]
        public string Process { get; set; }
        [Display(Name = "Vintage")]
        public string Vintage { get; set; }
        [Display(Name = "Wafer Size")]
        public string WaferSize { get; set; }
        [Display(Name = "Serial No")]
        public string SerialNo { get; set; }
        [Display(Name = "Config")]
        public string Config { get; set; }
        [Display(Name = "Fab")]
        public string Fab { get; set; }
        [Display(Name = "Code")]
        public string Code { get; set; }
        [Display(Name = "Location")]
        public string Location { get; set; }
        [Display(Name = "Inspector")]
        public string Inspector { get; set; }
        [Display(Name = "Inspection Summary")]
        public string InspectionSummary { get; set; }
        [Display(Name = "Remark")]
        public string Remark { get; set; }
        [Display(Name = "Period")]
        public string Period { get; set; }
        [Display(Name = "*Buy Date")]
        public string BuyDate { get; set; }
        [Display(Name = "*Sell Date")]
        public string SellDate { get; set; }
        [Display(Name = "Buyer")]
        public string Buyer { get; set; }
        [Display(Name = "Seller")]
        public string Seller { get; set; }
        [Display(Name = "Tool Price(B)")]
        public string ToolPrice_B { get; set; }
        [Display(Name = "Total Cost(B)")]
        public string TotalCost_B { get; set; }
        [Display(Name = "SG Cost(B)")]
        public string SGCostB { get; set; }
        [Display(Name = "Total Buy")]
        public string TotalBuy { get; set; }
        [Display(Name = "SG Total Buy")]
        public string SGTotalBuy { get; set; }
        [Display(Name = "*Sell Price(E)")]
        public string SellPrice_E { get; set; }
        [Display(Name = "*Target Price")]
        public string TargetPrice { get; set; }
        [Display(Name = "Profit")]
        public string Profit { get; set; }
        [Display(Name = "ROI")]
        public string ROI { get; set; }
        [Display(Name = "Annual ROI")]
        public string AnnualROI { get; set; }
        [Display(Name = "DeinstallCost(B)")]
        public string DeinstallCost_B { get; set; }
        [Display(Name = "Rigging Cost(B)")]
        public string RiggingCost_B { get; set; }
        [Display(Name = "Shipping Cost(B)")]
        public string ShippingCost_B { get; set; }
        [Display(Name = "Packing Cost(B)")]
        public string PackingCost_B { get; set; }
        [Display(Name = "Inland Trucking Cost(B)")]
        public string InlandTruckingCost_B { get; set; }
        [Display(Name = "Commission(B)")]
        public string Commission_B { get; set; }
        [Display(Name = "Warehouse Cost")]
        public string WarehouseCost { get; set; }
        [Display(Name = "SG Interest")]
        public string SGInterest { get; set; }
        [Display(Name = "Inventory allowance")]
        public string InventoryAllowance { get; set; }
        [Display(Name = "SG Commission")]
        public string SGCommission { get; set; }
        [Display(Name = "Task")]
        public string Task { get; set; }
        [Display(Name = "SG Offer USD")]
        public string SGOfferUSD { get; set; }

        //Extend columns
        [Display(Name = "Ext1")]
        public string Ext1 { get; set; }
        [Display(Name = "Ext2")]
        public string Ext2 { get; set; }
        [Display(Name = "Ext3")]
        public string Ext3 { get; set; }
        [Display(Name = "Ext4")]
        public string Ext4 { get; set; }
        [Display(Name = "Ext5")]
        public string Ext5 { get; set; }
        [Display(Name = "Ext6")]
        public string Ext6 { get; set; }
        [Display(Name = "Ext7")]
        public string Ext7 { get; set; }
        [Display(Name = "Ext8")]
        public string Ext8 { get; set; }
        [Display(Name = "Ext9")]
        public string Ext9 { get; set; }
        [Display(Name = "Ext10")]
        public string Ext10 { get; set; }
        [Display(Name = "Ext11")]
        public string Ext11 { get; set; }
        [Display(Name = "Ext12")]
        public string Ext12 { get; set; }
        [Display(Name = "Ext13")]
        public string Ext13 { get; set; }
        [Display(Name = "Ext14")]
        public string Ext14 { get; set; }
        [Display(Name = "Ext15")]
        public string Ext15 { get; set; }
        [Display(Name = "Ext16")]
        public string Ext16 { get; set; }
        [Display(Name = "Ext17")]
        public string Ext17 { get; set; }
        [Display(Name = "Ext18")]
        public string Ext18 { get; set; }
        [Display(Name = "Ext19")]
        public string Ext19 { get; set; }
        [Display(Name = "Ext20")]
        public string Ext20 { get; set; }
        [Display(Name = "Ext21")]
        public string Ext21 { get; set; }
        [Display(Name = "Ext22")]
        public string Ext22 { get; set; }
        [Display(Name = "Ext23")]
        public string Ext23 { get; set; }
        [Display(Name = "Ext24")]
        public string Ext24 { get; set; }
        [Display(Name = "Ext25")]
        public string Ext250 { get; set; }
        [Display(Name = "Ext26")]
        public string Ext26 { get; set; }
        [Display(Name = "Ext27")]
        public string Ext27 { get; set; }
        [Display(Name = "Ext28")]
        public string Ext28 { get; set; }
        [Display(Name = "Ext29")]
        public string Ext29 { get; set; }
        [Display(Name = "Ext30")]
        public string Ext30 { get; set; }

    }
}