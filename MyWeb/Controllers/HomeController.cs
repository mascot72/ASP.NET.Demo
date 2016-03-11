﻿using Excel;
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
        public ActionResult Upload(HttpPostedFileBase upload)
        {
            if (upload != null && upload.FileName != null)
                return UploadFirst(upload);

            string xlsPath = @"C:\workspace\resource\Cleansing (1st)";
            var dir = new System.IO.DirectoryInfo(xlsPath);
            var fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories);
            var files = from file in fileList
                        where file.Extension.Contains("xls")
                        select file;
            var fileName = files.FirstOrDefault().FullName;

            if (ModelState.IsValid)
            {
                ExcelConverter proc = new ExcelConverter();

                //DataSet ds = proc.OfficeExcelTODataSet(fileName);
                DataSet ds = proc.OleDBExcelToDataSet(fileName);

                return View(ds.Tables[1]);
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