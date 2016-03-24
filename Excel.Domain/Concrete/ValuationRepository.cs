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
        public bool AddModel(Valuation model)
        {
            bool result = default(bool);

            try
            {
                //model.CreateDate = DateTime.Now;
                this.context.Valuations.Add(model);
                result = this.context.SaveChanges() > 0;
                if (model.ExtendContent != null)
                {
                    foreach (var extCont in model.ExtendContent)
                    {
                        extCont.ImportID = model.ID;
                    }
                }
            }
            catch (Exception ex)
            {
                this.log.Error(MethodBase.GetCurrentMethod().Name, ex);
                throw ex;
            }

            return result;
        }

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

        public bool UpdateCleaning()
        {
            bool result = default(bool);

            try
            {
                result = context.Database.ExecuteSqlCommand(@"
update Valuations
set BuyDate = null
where convert(varchar(10), BuyDate, 126) = '1900-01-01'

update Valuations
set SellDate = null
where convert(varchar(10), SellDate, 126) = '1900-01-01'

update Valuations
set Date = null
where convert(varchar(10), Date, 126) = '1900-01-01'") > 0;

                return result;
            }
            catch (Exception ex)
            {
                this.log.Error(MethodBase.GetCurrentMethod().Name, ex);
                throw ex;
            }
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
