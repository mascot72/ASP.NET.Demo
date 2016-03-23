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
    public class ValuationRepository : LogBase, IDisposable, IValuationRepository
    {
        private EFDbContext context;

        /// <summary>
        /// Constractor
        /// </summary>
        public ValuationRepository() : base()
        {
            this.context = new EFDbContext();
        }

        //GetSingle

        //GetList

        /// <summary>
        /// GetList
        /// </summary>
        public IEnumerable<Valuation> Valuations
        {
            get
            {
                return context.Valuations;
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
                //추가정보 삭제
                ExtendContentRepository rep = new ExtendContentRepository();

                var single = this.context.Valuations.Find(Id);
                if (single != null)
                {
                    rep.RemoveModelByParent(Id);    //자식삭제
                    this.context.Valuations.Remove(single); //자신삭제
                    result = this.context.SaveChanges() > 0;
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
