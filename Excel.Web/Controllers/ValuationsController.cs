using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Excel.Domain.Concrete;
using Excel.Domain.Entites;

namespace Excel.Web.Controllers
{
    public class ValuationsController : Controller
    {
        private EFDbContext db = new EFDbContext();

        // GET: Valuations
        public ActionResult Index()
        {
            var valuations = db.Valuations.Include(v => v.FileImport);
            return View(valuations.ToList());
        }

        // GET: Valuations/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Valuation valuation = db.Valuations.Find(id);
            if (valuation == null)
            {
                return HttpNotFound();
            }
            return View(valuation);
        }

        // GET: Valuations/Create
        public ActionResult Create()
        {
            ViewBag.FileID = new SelectList(db.FileImports, "ID", "Path");
            return View();
        }

        // POST: Valuations/Create
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,FileID,EID,InvenNo,SGNo,TID,Date,Name,Version,Type,DealNo,LeadNo,Comment,Comment_1,Currency,Category,Maker,Model,Process,Vintage,WaferSize,SerialNo,Config,Fab,Code,Location,Inspector,InspectionSummary,Remark,Period,BuyDate,SellDate,Buyer,Seller,ToolPriceB,TotalCostB,SGCostB,TotalCostS,TotalBuy,SGTotalBuy,SellPriceE,TargetPrice,Profit,ProfitPercent,ROI,AnnualROI,DeinstallCostB,RiggingCostB,ShippingCostB,PackingCostB,InlandTruckingCostB,CommissionB,WarehouseCost,SGWarehouseCost,SGInterest,InventoryAllowance,SGCommission,Task,SGOfferUSD,Qty,Ext1,Ext2,Ext3,Ext4,Ext5,Ext6,Ext7,Ext8,Ext9,Ext10,Ext11,Ext12,Ext13,Ext14,Ext15,Ext16,Ext17,Ext18,Ext19,Ext20,Ext21,Ext22,Ext23,Ext24,Ext25,Ext26,Ext27,Ext28,Ext29,Ext30,Ext31,Ext32,Ext33,Ext34,Ext35,Ext36,Ext37,Ext38,Ext39,Ext40,Ext41,Ext42,Ext43,Ext44,Ext45,Ext46,Ext47,Ext48,Ext49,Ext50,Ref1,Ref2,Reason,Creator")] Valuation valuation)
        {
            if (ModelState.IsValid)
            {
                db.Valuations.Add(valuation);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            ViewBag.FileID = new SelectList(db.FileImports, "ID", "Path", valuation.FileID);
            return View(valuation);
        }

        // GET: Valuations/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Valuation valuation = db.Valuations.Find(id);
            if (valuation == null)
            {
                return HttpNotFound();
            }
            ViewBag.FileID = new SelectList(db.FileImports, "ID", "Path", valuation.FileID);
            return View(valuation);
        }

        // POST: Valuations/Edit/5
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,FileID,EID,InvenNo,SGNo,TID,Date,Name,Version,Type,DealNo,LeadNo,Comment,Comment_1,Currency,Category,Maker,Model,Process,Vintage,WaferSize,SerialNo,Config,Fab,Code,Location,Inspector,InspectionSummary,Remark,Period,BuyDate,SellDate,Buyer,Seller,ToolPriceB,TotalCostB,SGCostB,TotalCostS,TotalBuy,SGTotalBuy,SellPriceE,TargetPrice,Profit,ProfitPercent,ROI,AnnualROI,DeinstallCostB,RiggingCostB,ShippingCostB,PackingCostB,InlandTruckingCostB,CommissionB,WarehouseCost,SGWarehouseCost,SGInterest,InventoryAllowance,SGCommission,Task,SGOfferUSD,Qty,Ext1,Ext2,Ext3,Ext4,Ext5,Ext6,Ext7,Ext8,Ext9,Ext10,Ext11,Ext12,Ext13,Ext14,Ext15,Ext16,Ext17,Ext18,Ext19,Ext20,Ext21,Ext22,Ext23,Ext24,Ext25,Ext26,Ext27,Ext28,Ext29,Ext30,Ext31,Ext32,Ext33,Ext34,Ext35,Ext36,Ext37,Ext38,Ext39,Ext40,Ext41,Ext42,Ext43,Ext44,Ext45,Ext46,Ext47,Ext48,Ext49,Ext50,Ref1,Ref2,Reason,Creator")] Valuation valuation)
        {
            if (ModelState.IsValid)
            {
                db.Entry(valuation).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            ViewBag.FileID = new SelectList(db.FileImports, "ID", "Path", valuation.FileID);
            return View(valuation);
        }

        // GET: Valuations/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Valuation valuation = db.Valuations.Find(id);
            if (valuation == null)
            {
                return HttpNotFound();
            }
            return View(valuation);
        }

        // POST: Valuations/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Valuation valuation = db.Valuations.Find(id);
            db.Valuations.Remove(valuation);
            db.SaveChanges();
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
