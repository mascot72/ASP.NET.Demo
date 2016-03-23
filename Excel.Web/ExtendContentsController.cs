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
    public class ExtendContentsController : ApiController
    {
        private EFDbContext db = new EFDbContext();

        // GET: api/ExtendContents
        public IQueryable<ExtendContent> GetExtendContents()
        {
            return db.ExtendContents;
        }

        // GET: api/ExtendContents/5
        [ResponseType(typeof(ExtendContent))]
        public IHttpActionResult GetExtendContent(int id)
        {
            ExtendContent extendContent = db.ExtendContents.Find(id);
            if (extendContent == null)
            {
                return NotFound();
            }

            return Ok(extendContent);
        }

        // PUT: api/ExtendContents/5
        [ResponseType(typeof(void))]
        public IHttpActionResult PutExtendContent(int id, ExtendContent extendContent)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != extendContent.ID)
            {
                return BadRequest();
            }

            db.Entry(extendContent).State = EntityState.Modified;

            try
            {
                db.SaveChanges();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ExtendContentExists(id))
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

        // POST: api/ExtendContents
        [ResponseType(typeof(ExtendContent))]
        public IHttpActionResult PostExtendContent(ExtendContent extendContent)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.ExtendContents.Add(extendContent);
            db.SaveChanges();

            return CreatedAtRoute("DefaultApi", new { id = extendContent.ID }, extendContent);
        }

        // DELETE: api/ExtendContents/5
        [ResponseType(typeof(ExtendContent))]
        public IHttpActionResult DeleteExtendContent(int id)
        {
            ExtendContent extendContent = db.ExtendContents.Find(id);
            if (extendContent == null)
            {
                return NotFound();
            }

            db.ExtendContents.Remove(extendContent);
            db.SaveChanges();

            return Ok(extendContent);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool ExtendContentExists(int id)
        {
            return db.ExtendContents.Count(e => e.ID == id) > 0;
        }
    }
}