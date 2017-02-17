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

        // GET: Process Rate
        public ActionResult GetProcessRate(int workCount = 10, string processState = "", string startDate = "")
        {
            DateTime stDate = startDate != null && !string.IsNullOrEmpty(startDate) ? DateTime.Parse(startDate) : DateTime.MinValue;
            var resultVal = new { targetCount = db.FileImports.Count(), successfulCount = db.FileImports.Where(e => e.CreateDate >= stDate).Count() };
            return Json(resultVal, JsonRequestBehavior.AllowGet);
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

        /// <summary>
        /// Excel Import Processing
        /// </summary>
        /// <param name="upload"></param>
        /// <param name="isReadonly">화면으로 조회만 할때</param>
        /// <param name="workCount">작업대상 파일수</param>
        /// <param name="processState">처리할 상태</param>
        /// <param name="folderPath">폴더경로</param>
        /// <returns></returns>
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
            }

            try
            {
                string xlsPath = @"";
                if (folderPath != "")
                {
                    xlsPath = folderPath;
                }
                else
                {
                    new ApplicationException("파일경로가 없습니다");
                }

                var dir = new System.IO.DirectoryInfo(xlsPath);
                IEnumerable<FileInfo> files = null;
                if (dir.Exists)
                {
                    var fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories);
                    files = from file in fileList
                            where file.Extension.ToLower().Contains("xls")
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
                        targetFiles = files.Count();

                        List<string> extCol = new List<string>();
                        var fileTable = proc.GetFileTable();

                        if (processState != "")
                        {
                            fileTable = fileTable.Where(e => e.Result == processState).ToList();
                        }

                        int[] result = { 0, 0, 0 };

                        if (fileName == string.Empty)
                        {
                            IEnumerable<FileInfo> lst;
                            if (processState != "")
                            {
                                //이력이 있는 파일들
                                lst = from aa in files
                                      join bb in fileTable on aa.Name equals bb.Name
                                      where bb.Result.Equals(processState)
                                      select aa;
                            }
                            else
                            {
                                //lst = from aa in files
                                //      where !fileTable.Select(y=>y.Name).Equals(aa.Name)
                                //      select aa;

                                var groupBNames = new HashSet<string>(fileTable.Select(x => x.Name));
                                lst = files.Where(x => !groupBNames.Contains(x.Name));

                                //lst = files.Where(x => !fileTable.Select(b => b.Name).Contains(x.Name));  //이력이 없는 파일들
                            }

                            foreach (var file in lst)
                            {
                                try
                                {
                                    if (workCount <= 0) break;

                                    ds = proc.ExcelToDB(file.FullName, result, processState);
                                    if (ds != null)
                                    {
                                        resultFiles += result[0];
                                        resultRows += result[1];
                                        targetRowCount += result[2];
                                        targetFiles++;
                                    }
                                    workCount--;    //처리된 만큼 처리할 행수 뺀다

                                }
                                catch (Exception ex)
                                {
                                    ViewBag.Message = ex.Message;
                                }
                            }
                        }
                        else
                        {
                            bool pass = false;
                            if (processState != "")
                            {
                                pass = fileTable.Where(b => b.Path + "\\" + b.Name == fileName && b.Result == processState).Count() > 0;   //이력이 있는 파일
                            }
                            else
                            {
                                pass = fileTable.Where(b => b.Path + "\\" + b.Name == fileName).Count() == 0;  //이력이 없는 파일들
                            }
                            if (pass)
                            {
                                targetFiles = 1;
                                ds = proc.ExcelToDB(fileName, result, processState);
                                resultFiles += result[0];
                                resultRows += result[1];
                                targetRowCount += result[2];
                            }
                        }

                        //DB 후처리 작업수행
                        if (targetRowCount > 0)
                        {
                            using (ValuationRepository mstContext = new ValuationRepository())
                            {
                                mstContext.UpdateCleaning();
                            }
                        }

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

        //1개의 파일 Import
        [HttpPost]
        public ActionResult MyFileUpload(string processState = "", string filePath = "")
        {
            var fileName = string.Empty;
            int targetFiles = 0;
            int targetRowCount = 0;
            DataSet ds = null;
            bool pass = false;

            fileName = filePath;            

            try
            {
                if (string.IsNullOrEmpty(fileName)) return Json(null);

                var file = new System.IO.FileInfo(fileName);
                if (file.Exists)
                {
                    if (ModelState.IsValid)
                    {
                        Excel.Web.Processor.ExcelConverter proc = new Excel.Web.Processor.ExcelConverter();

                      
                            int resultFiles = 0;
                            int resultRows = 0;
                            int[] result = { 0, 0, 0 };

                            List<string> extCol = new List<string>();
                            var fileTable = proc.GetFileTable();

                            if (processState != "")
                            {
                                fileTable = fileTable.Where(e => e.Result == processState).ToList();
                            }
                            
                            if (processState != "")
                            {
                                pass = fileTable.Where(b => b.Path + "\\" + b.Name == fileName && b.Result == processState).Count() > 0;   //이력이 있는 파일
                            }
                            else
                            {
                                pass = fileTable.Where(b => b.Path + "\\" + b.Name == fileName).Count() == 0;  //이력이 없는 파일들
                            }
                            if (pass)
                            {
                                targetFiles = 1;
                                ds = proc.ExcelToDB(fileName, result, processState);
                            }

                            //DB 후처리 작업수행
                            if (targetRowCount > 0)
                            {
                                using (ValuationRepository mstContext = new ValuationRepository())
                                {
                                    mstContext.UpdateCleaning();
                                }
                            }

                            ViewBag.Message = string.Format("Success File Count({0}/{1}) \r\nSuccess Row Count({2}/{3})", resultFiles, targetFiles, resultRows, targetRowCount);
                        
                    }

                    //결과반환
                    if (ds != null && ds.Tables.Count == 2)
                        return Json(1);
                    else
                        return Json(0);
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                return Json(0);
            }

            return Json(0);
        }

        //파일정보가져오기
        [HttpPost]
        public ActionResult GetUploadableData(HttpPostedFileBase upload, string isReadonly, int workCount = 10, string processState = "", string folderPath = "")
        {
            var fileName = string.Empty;
            FileInfo[] fileList = null;
            IEnumerable<FileInfo> files1 = null;
            int targetFiles = 0;

            if (upload != null && upload.FileName != null)
            {
                fileName = upload.FileName;
            }

            try
            {
                string xlsPath = @"";
                if (folderPath != "")
                {
                    xlsPath = folderPath;
                }
                else
                {
                    new ApplicationException("파일경로가 없습니다");
                }

                var dir = new System.IO.DirectoryInfo(xlsPath);

                if (dir.Exists)
                {
                    fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories);
                    var files = from item in fileList
                                where item.Extension.ToLower().Contains("xls")
                                select new { Name = item.Name, Length = item.Length, Directory = item.Directory.FullName, selected = false };

                    releaseObject(fileList);
                    fileList = null;

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
                            targetFiles = files.Count();

                            List<string> extCol = new List<string>();
                            var fileTable = proc.GetFileTable();

                            if (processState != "")
                            {
                                fileTable = fileTable.Where(e => e.Result == processState).ToList();
                            }

                            IEnumerable<FileInfo> lst;
                            if (processState != "") //처리했던 파일중 완료(S)|에러(E)인 상태 조회
                            {
                                //이력이 있는 파일들
                                var result = from aa in files
                                           join bb in fileTable on aa.Name equals bb.Name
                                           where bb.Result.Equals(processState)
                                           select new { Name = aa.Name, Length = aa.Length, Directory = aa.Directory, selected = false };

                                return Json(result);
                            }
                            else
                            {
                                //DB에 없는 파일(처리하지 않은)의 목록
                                var groupBNames = new HashSet<string>(fileTable.Select(x => x.Name));

                                var listFile = fileTable.Select(e => new { Name = e.Name, Length = (long)e.Size, Directory = e.Path, selected = false });
                                //var listDbFile = fileTable.Select(e => new { Name = e.Name, selected = false });
                                var result = files.Except(listFile);    //차(file목록 - db처리한목록), 처리한 모든것을 빼기때문에 에러난 것도 제외된다

                                return Json(result);
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
            }
            finally
            {

            }
            return Json("");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                //GC.Collect();
            }
        }

        //화면에 보여주기만 한다.
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

                    if (upload.FileName.ToLower().EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (upload.FileName.ToLower().EndsWith(".xlsx"))
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
