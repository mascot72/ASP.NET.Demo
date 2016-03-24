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
using System.IO;

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

        public ActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Upload(HttpPostedFileBase upload, string isReadonly, int workCount = 10, string processState = "", string folderPath = "")
        {
            var fileName = string.Empty;
            int targetFiles = 0;
            int targetRowCount = 0;

            if (upload != null && upload.FileName != null)
            {
                fileName = upload.FileName;
                //return UploadFirst(upload);
            }

            try
            {
                string xlsPath = @"C:\workspace\resource\Cleansing (1st)\New Valuation";
                if (folderPath != "") {
                    xlsPath = folderPath;
                }

                var dir = new System.IO.DirectoryInfo(xlsPath);
                IEnumerable<FileInfo> files = null;
                if (dir.Exists)
                {
                    var fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories);
                    files = from file in fileList
                            where file.Extension.Contains("xls")
                            select file;
                }

                if (ModelState.IsValid)
                {
                    Excel.Web.Processor.ExcelConverter proc = new Excel.Web.Processor.ExcelConverter();
                    DataSet ds = null;
                    if (isReadonly != null && isReadonly == "true")
                    {
                        return UploadFirst(upload);
                    }
                    else
                    {
                        int resultFiles = 0;
                        int resultRows = 0;
                        List<string> extCol = new List<string>();
                        var fileTable = proc.GetFileTable();
                        int[] result = { 0, 0, 0 };

                        if (fileName == string.Empty)
                        {
                            foreach (var file in files)
                            {
                                try
                                {
                                    //Error항목처리시
                                    string filefullName = file.FullName;

                                    //오류만처리(임시)
                                    /*
                                    if (fileTable.Where(e => e.Name == file.Name && e.Result == "E").Count() > 0)
                                    {
                                        if (workCount <= 0) break;
                                        ds = proc.ExcelToDB(filefullName, result);
                                        workCount--;    //처리된 만큼 처리할 행수 뺀다
                                        resultFiles += ds.Tables[0].Rows.Count;
                                        resultRows += ds.Tables[1].Rows.Count;
                                        targetFiles++;
                                    }
                                    else
                                        continue;
                                    */

                                    if (processState != "")
                                    {
                                        if (fileTable.Where(e => e.Result == processState && e.Name == file.Name).Count() == 0)  //오류났던 항목이면 처리(결과가 상태와 같고 처리된 파일명이 없으면 제외)
                                        {
                                            continue;
                                        }
                                    }
                                    else if (fileTable.Where(e => e.Name == file.Name).Count() > 0)
                                    {
                                        continue;   //완료된 파일은 제외!(Error완료항목 요청이 아니고 기존에 존재시 통과)
                                    }

                                    if (workCount <= 0) break;
                                    ds = proc.ExcelToDB(filefullName, result);
                                    workCount--;    //처리된 만큼 처리할 행수 뺀다
                                    resultFiles += result[0];
                                    resultRows += result[1];
                                    targetRowCount += result[2];
                                    targetFiles++;
                                }
                                catch (Exception ex)
                                {
                                    //throw;
                                    ViewBag.Message = ex.Message;
                                }
                            }
                        }
                        else
                        {
                            ds = proc.ExcelToDB(fileName, result);
                        }
                        //proc.UpdateAfter(); //DB 후처리 작업수행
                        //using (ValuationRepository mstContext = new ValuationRepository())
                        //{
                        //    mstContext.UpdateCleaning();
                        //}

                        ViewBag.Message = string.Format("Success File Count({0}/{1}) \r\nSuccess Row Count({2}/{3})", resultFiles, targetFiles, resultRows, targetRowCount);
                    }

                    if (ds != null && ds.Tables.Count == 2)
                        return View(ds.Tables[1]);
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                return View();
            }



            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult UploadFirst(HttpPostedFileBase upload)
        {
            if (ModelState.IsValid)
            {
                //Install-Package ExcelDataReader
                if (upload != null && upload.ContentLength > 0)
                {
                    // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                    // to get started. This is how we avoid dependencies on ACE or Interop:
                    Stream stream = upload.InputStream;

                    // We return the interface, so that
                    IExcelDataReader reader = null;

                    if (upload.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (upload.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");
                        return View();
                    }

                    reader.IsFirstRowAsColumnNames = true;

                    DataSet result = reader.AsDataSet();
                    reader.Close();

                    return View(result.Tables[0]);
                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }
            return View();
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
