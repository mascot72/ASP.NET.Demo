using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MyWeb.Processor;

namespace MyWeb.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
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
            int targetRows = 0;

            if (upload != null && upload.FileName != null)
            {
                fileName = upload.FileName;
                //return UploadFirst(upload);
            }

            try
            {
                string xlsPath = @"C:\workspace\resource\Cleansing (1st)\New Valuation";
                if (folderPath != "") xlsPath = folderPath;

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
                    ExcelConverter proc = new ExcelConverter();
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
                        int[] result = { 0, 0 };

                        if (fileName == string.Empty)
                        {
                            foreach (var file in files)
                            {
                                try
                                {
                                    //Error항목처리시
                                    string filefullName = file.FullName;
                                    
                                    if (processState != "")
                                    {
                                        if (fileTable.Where(e => e.Result == processState && e.Name == file.Name).Count() == 0)  //오류났던 항목이면 처리
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
                                    workCount--;
                                    resultFiles += ds.Tables[0].Rows.Count;
                                    resultRows += ds.Tables[1].Rows.Count;
                                    targetFiles++;
                                }
                                catch (Exception)
                                {
                                    //throw;
                                }
                            }
                        }
                        else
                        {
                            ds = proc.ExcelToDB(fileName, result);
                        }
                        proc.UpdateAfter(); //DB ㄷ후처리 작업수행
                        ViewBag.Message = string.Format("Success File Count({0}/{1}) \r\nSuccess Row Count({2}/{3})", result[0], targetFiles, result[1], resultRows);
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
    }
}