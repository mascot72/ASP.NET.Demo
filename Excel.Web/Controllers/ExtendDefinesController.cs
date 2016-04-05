using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using Excel.Domain.Concrete;
using Excel.Domain.Entites;

namespace Excel.Web.Views.Valuations
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
        public async Task<IHttpActionResult> GetExtendDefine(int id)
        {
            ExtendDefine extendDefine = await db.ExtendDefines.FindAsync(id);
            if (extendDefine == null)
            {
                return NotFound();
            }

            return Ok(extendDefine);
        }

        // PUT: api/ExtendDefines/5
        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutExtendDefine(int id, ExtendDefine extendDefine)
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
                await db.SaveChangesAsync();
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
        public async Task<IHttpActionResult> PostExtendDefine(ExtendDefine extendDefine)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            db.ExtendDefines.Add(extendDefine);
            await db.SaveChangesAsync();

            return CreatedAtRoute("DefaultApi", new { id = extendDefine.ID }, extendDefine);
        }

        // DELETE: api/ExtendDefines/5
        [ResponseType(typeof(ExtendDefine))]
        public async Task<IHttpActionResult> DeleteExtendDefine(int id)
        {
            ExtendDefine extendDefine = await db.ExtendDefines.FindAsync(id);
            if (extendDefine == null)
            {
                return NotFound();
            }

            db.ExtendDefines.Remove(extendDefine);
            await db.SaveChangesAsync();

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