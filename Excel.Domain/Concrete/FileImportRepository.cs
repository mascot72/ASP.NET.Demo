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
    public class FileImportRepository : LogBase, IDisposable, IFileImportRepository
    {
        private EFDbContext context;

        /// <summary>
        /// Constractor
        /// </summary>
        public FileImportRepository() : base()
        {
            this.context = new EFDbContext();
        }

        //GetSingle

        //GetList
        public IEnumerable<FileImport> FileImports
        {
            get
            {
                return context.FileImports;
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
