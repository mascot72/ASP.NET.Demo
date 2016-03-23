using Excel.Domain.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel.Domain.Entites;
using System.Reflection;

namespace Excel.Domain.Concrete
{
    public class ExtendContentRepository : LogBase, IDisposable, IExtendContentRepository
    {
        private EFDbContext context;

        /// <summary>
        /// Constractor
        /// </summary>
        public ExtendContentRepository() : base()
        {
            this.context = new EFDbContext();
        }

        //GetSingle

        //GetList
        public IEnumerable<ExtendContent> ExtendContents
        {
            get
            {
                return context.ExtendContents;
            }
        }

        //Create

        //Remove

        /// <summary>
        /// Remove
        /// </summary>
        /// <param name="Id">Valuation ID</param>
        /// <returns></returns>
        public bool RemoveModel(int Id)
        {
            bool result = default(bool);

            try
            {
                var extCont = this.context.ExtendContents.Find(Id);

                if (extCont != null)
                {
                    this.context.ExtendContents.Attach(extCont);
                    var entry = this.context.Entry(extCont);
                    return this.context.SaveChanges() > 0;
                }
            }
            catch (Exception ex)
            {
                this.log.Error(MethodBase.GetCurrentMethod().Name, ex);
            }

            return result;
        }

        public bool RemoveModelByParent(int parentId)
        {
            bool result = default(bool);

            try
            {
                result = this.context.ExtendContents.ToList().RemoveAll(x => x.ImportID == parentId) > 0;
                
            }
            catch (Exception ex)
            {
                this.log.Error(MethodBase.GetCurrentMethod().Name, ex);
            }

            return result;
        }

        /// <summary>
        /// 소멸자
        /// </summary>
        public void Dispose()
        {
            this.context.Dispose();
            GC.Collect();
        }
    }
}
