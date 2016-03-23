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
    public class FileImportsController : Controller
    {
        private EFDbContext db = new EFDbContext();

        // GET: FileImports
        public ActionResult Index()
        {
            return View(db.FileImports.ToList());
        }

        // GET: FileImports/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FileImport fileImport = db.FileImports.Find(id);
            if (fileImport == null)
            {
                return HttpNotFound();
            }
            return View(fileImport);
        }

        // GET: FileImports/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: FileImports/Create
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Path,Name,ExtName,Result,Reason,Remark,Extend,CreateDate,Creator,Size")] FileImport fileImport)
        {
            if (ModelState.IsValid)
            {
                db.FileImports.Add(fileImport);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(fileImport);
        }

        // GET: FileImports/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FileImport fileImport = db.FileImports.Find(id);
            if (fileImport == null)
            {
                return HttpNotFound();
            }
            return View(fileImport);
        }

        // POST: FileImports/Edit/5
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Path,Name,ExtName,Result,Reason,Remark,Extend,CreateDate,Creator,Size")] FileImport fileImport)
        {
            if (ModelState.IsValid)
            {
                db.Entry(fileImport).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(fileImport);
        }

        // GET: FileImports/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            FileImport fileImport = db.FileImports.Find(id);
            if (fileImport == null)
            {
                return HttpNotFound();
            }
            return View(fileImport);
        }

        // POST: FileImports/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            FileImport fileImport = db.FileImports.Find(id);
            db.FileImports.Remove(fileImport);
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
