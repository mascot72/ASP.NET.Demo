using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MyWeb.Models;
using MyWeb.Models.Excel;

namespace MyWeb.Controllers
{
    public class ValuationModelsController : Controller
    {
        private ApplicationDbContext db = new ApplicationDbContext();

        // GET: ValuationModels
        public async Task<ActionResult> Index()
        {
            return View(await db.ValuationModels.ToListAsync());
        }

        // GET: ValuationModels/Details/5
        public async Task<ActionResult> Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ValuationModels valuationModels = await db.ValuationModels.FindAsync(id);
            if (valuationModels == null)
            {
                return HttpNotFound();
            }
            return View(valuationModels);
        }

        // GET: ValuationModels/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ValuationModels/Create
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "ID,EId,InvenNo,SGNo,TId,Date,Name,Version,Type,DealNo,LeadNo,Comment,Currency,Category,Maker,Model,Process,Vintage,WaferSize,SerialNo,Config,Fab,Code,Location,Inspector,InspectionSummary,Remark,Period,BuyDate,SellDate,Buyer,Seller,ToolPrice_B,TotalCost_B,SGCostB,TotalBuy,SGTotalBuy,SellPrice_E,TargetPrice,Profit,ROI,AnnualROI,DeinstallCost_B,RiggingCost_B,ShippingCost_B,PackingCost_B,InlandTruckingCost_B,Commission_B,WarehouseCost,SGInterest,InventoryAllowance,SGCommission,Task,SGOfferUSD,Ext1,Ext2,Ext3,Ext4,Ext5,Ext6,Ext7,Ext8,Ext9,Ext10,Ext11,Ext12,Ext13,Ext14,Ext15,Ext16,Ext17,Ext18,Ext19,Ext20,Ext21,Ext22,Ext23,Ext24,Ext250,Ext26,Ext27,Ext28,Ext29,Ext30")] ValuationModels valuationModels)
        {
            if (ModelState.IsValid)
            {
                db.ValuationModels.Add(valuationModels);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            return View(valuationModels);
        }

        // GET: ValuationModels/Edit/5
        public async Task<ActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ValuationModels valuationModels = await db.ValuationModels.FindAsync(id);
            if (valuationModels == null)
            {
                return HttpNotFound();
            }
            return View(valuationModels);
        }

        // POST: ValuationModels/Edit/5
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "ID,EId,InvenNo,SGNo,TId,Date,Name,Version,Type,DealNo,LeadNo,Comment,Currency,Category,Maker,Model,Process,Vintage,WaferSize,SerialNo,Config,Fab,Code,Location,Inspector,InspectionSummary,Remark,Period,BuyDate,SellDate,Buyer,Seller,ToolPrice_B,TotalCost_B,SGCostB,TotalBuy,SGTotalBuy,SellPrice_E,TargetPrice,Profit,ROI,AnnualROI,DeinstallCost_B,RiggingCost_B,ShippingCost_B,PackingCost_B,InlandTruckingCost_B,Commission_B,WarehouseCost,SGInterest,InventoryAllowance,SGCommission,Task,SGOfferUSD,Ext1,Ext2,Ext3,Ext4,Ext5,Ext6,Ext7,Ext8,Ext9,Ext10,Ext11,Ext12,Ext13,Ext14,Ext15,Ext16,Ext17,Ext18,Ext19,Ext20,Ext21,Ext22,Ext23,Ext24,Ext250,Ext26,Ext27,Ext28,Ext29,Ext30")] ValuationModels valuationModels)
        {
            if (ModelState.IsValid)
            {
                db.Entry(valuationModels).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(valuationModels);
        }

        // GET: ValuationModels/Delete/5
        public async Task<ActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ValuationModels valuationModels = await db.ValuationModels.FindAsync(id);
            if (valuationModels == null)
            {
                return HttpNotFound();
            }
            return View(valuationModels);
        }

        // POST: ValuationModels/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(int id)
        {
            ValuationModels valuationModels = await db.ValuationModels.FindAsync(id);
            db.ValuationModels.Remove(valuationModels);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
