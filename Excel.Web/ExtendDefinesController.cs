using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Description;
using Excel.Domain.Concrete;
using Excel.Domain.Entites;

namespace Excel.Web
{
    public class ExtendDefinesController : ApiController
    {
        private EFDbContext db = new EFDbContext();

        // GET: api/ExtendDefines
        public IQueryable<ExtendDefine> GetExtendDefines()
        {
            return db.ExtendDefines;
        }

        // GET: api/ExtendDefines/5
        [ResponseType(typeof(ExtendDefine))]
        public IHttpActionResult GetExtendDefine(int id)
        {
            ExtendDefine extendDefine = db.ExtendDefines.Find(id);
            if (extendDefine == null)
            {
                return NotFound();
            }

            return Ok(extendDefine);
        }

        // PUT: api/ExtendDefines/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutExtendDefine(int id, ExtendDefine extendDefine)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != extendDefine.ID)
            {
                return BadRequest();
            }

            db.Entry(extendDefine).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ExtendDefineExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        // POST: api/ExtendDefines
        [ResponseType(typeof(ExtendDefine))]
        public IHttpActionResult PostExtendDefine(ExtendDefine extendDefine)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.ExtendDefines.Add(extendDefine);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = extendDefine.ID }, extendDefine);
        }

        // DELETE: api/ExtendDefines/5
        [ResponseType(typeof(ExtendDefine))]
        public IHttpActionResult DeleteExtendDefine(int id)
        {
            ExtendDefine extendDefine = db.ExtendDefines.Find(id);
            if (extendDefine == null)
            {
                return NotFound();
            }

            db.ExtendDefines.Remove(extendDefine);
            db.SaveChanges();

            return Ok(extendDefine);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool ExtendDefineExists(int id)
        {
            return db.ExtendDefines.Count(e => e.ID == id) > 0;
        }
    }
}