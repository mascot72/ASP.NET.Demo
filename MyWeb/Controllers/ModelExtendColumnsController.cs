using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MyWeb.Models;

namespace MyWeb.Controllers
{
    public class ModelExtendColumnsController : Controller
    {
        private ApplicationDbContext db = new ApplicationDbContext();

        // GET: ModelExtendColumns
        public ActionResult Index()
        {
            return View(db.ModelExtendColumns.ToList());
        }

        // GET: ModelExtendColumns/Details/5
        public ActionResult Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ModelExtendColumn modelExtendColumn = db.ModelExtendColumns.Find(id);
            if (modelExtendColumn == null)
            {
                return HttpNotFound();
            }
            return View(modelExtendColumn);
        }

        // GET: ModelExtendColumns/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ModelExtendColumns/Create
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "ID,Name,CreateDate")] ModelExtendColumn modelExtendColumn)
        {
            if (ModelState.IsValid)
            {
                db.ModelExtendColumns.Add(modelExtendColumn);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(modelExtendColumn);
        }

        // GET: ModelExtendColumns/Edit/5
        public ActionResult Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ModelExtendColumn modelExtendColumn = db.ModelExtendColumns.Find(id);
            if (modelExtendColumn == null)
            {
                return HttpNotFound();
            }
            return View(modelExtendColumn);
        }

        // POST: ModelExtendColumns/Edit/5
        // 초과 게시 공격으로부터 보호하려면 바인딩하려는 특정 속성을 사용하도록 설정하십시오. 
        // 자세한 내용은 http://go.microsoft.com/fwlink/?LinkId=317598을(를) 참조하십시오.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "ID,Name,CreateDate")] ModelExtendColumn modelExtendColumn)
        {
            if (ModelState.IsValid)
            {
                db.Entry(modelExtendColumn).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(modelExtendColumn);
        }

        // GET: ModelExtendColumns/Delete/5
        public ActionResult Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            ModelExtendColumn modelExtendColumn = db.ModelExtendColumns.Find(id);
            if (modelExtendColumn == null)
            {
                return HttpNotFound();
            }
            return View(modelExtendColumn);
        }

        // POST: ModelExtendColumns/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(string id)
        {
            ModelExtendColumn modelExtendColumn = db.ModelExtendColumns.Find(id);
            db.ModelExtendColumns.Remove(modelExtendColumn);
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
