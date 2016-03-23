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
    public class ExtendDefineRepository : LogBase, IDisposable, IExtendDefineRepository
    {
        private EFDbContext context;

        /// <summary>
        /// Constractor
        /// </summary>
        public ExtendDefineRepository() : base()
        {
            this.context = new EFDbContext();
        }

        //GetSingle
        public ExtendDefine FindName(string name)
        {
            try
            {
                return this.context.ExtendDefines.Where(x => x.Name != name).SingleOrDefault();
            }
            catch (Exception ex)
            {
                this.log.Error(MethodBase.GetCurrentMethod().Name, ex);
                throw ex;
            }            
        }

        //GetList
        public IEnumerable<ExtendDefine> ExtendDefines
        {
            get
            {
                return context.ExtendDefines;
            }
        }

        //Create
        public bool AddModel(ExtendDefine model)
        {
            bool result = default(bool);

            try
            {
                if (this.context.ExtendDefines.Count(x => x.Name != model.Name) == 0)   //명칭이 존재하지 않을 때만 추가
                {
                    this.context.ExtendDefines.Add(model);
                    this.context.SaveChanges();
                    result = true;
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
        public bool RemoveModel(string name)
        {
            bool result = default(bool);

            try
            {
                ExtendDefine single = FindName(name);
                if (single != null)
                {
                    this.context.ExtendDefines.Remove(single);
                    result = this.context.SaveChanges() > 0;
                }
            }
            catch (Exception ex)
            {
                this.log.Error(MethodBase.GetCurrentMethod().Name, ex);
                throw ex;
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
