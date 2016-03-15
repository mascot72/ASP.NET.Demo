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
        public ActionResult Upload(HttpPostedFileBase upload, string isReadonly)
        {
            var fileName = string.Empty;

            if (upload != null && upload.FileName != null)
            {
                fileName = upload.FileName;
                //return UploadFirst(upload);
            }

            try
            {
                string xlsPath = @"C:\workspace\resource\Cleansing (1st)";
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
                        int workCount = 100;

                        if (fileName == string.Empty)
                        {
                            foreach (var filefullName in (files.ToList().Select(e => e.FullName)))
                            {
                                if (workCount <= 0) break;
                                ds = proc.ExcelToDB(filefullName);
                                workCount--;
                            }
                        }
                        else
                        {
                            ds = proc.ExcelToDB(fileName);
                        }

                        ViewBag.Message = "Process File Count : " + ds.Tables[0].Rows.Count.ToString();
                        ViewBag.Message += "\r\nProcessed Row Count : " + ds.Tables[1].Rows.Count.ToString();                        
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