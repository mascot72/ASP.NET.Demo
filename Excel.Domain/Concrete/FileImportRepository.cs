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
        public FileImport GetModel(int Id)
        {
            try
            {
                return this.context.FileImports.Find(Id);
            }
            catch (Exception ex)
            {
                this.log.Error(MethodBase.GetCurrentMethod().Name, ex);
                throw ex;
            }
        }

        //GetList
        public IEnumerable<FileImport> FileImports
        {
            get
            {
                return context.FileImports;
            }
        }

        //Create

        //Update
        public bool ModifyModel(FileImport model)
        {
            bool result = default(bool);
            try
            {
                var original = this.context.FileImports.Find(model.ID);

                if (original != null)
                {                    
                    original.Result = model.Result;
                    original.Reason = model.Reason;
                    original.Remark = model.Remark;

                    //기존에 존재하는 파일은 비활성(내가 아니고 결과가 1개인것이고 파일명이 동일한 것)
                    var existData = this.context.FileImports.Where(x => x.Name == model.Name && x.Result.Length == 1 && x.ID != model.ID).ToList();
                    foreach(var item in existData)
                    {
                        item.Result += model.Result;
                        item.UpdateDate = DateTime.Now;
                        item.Updater = model.Updater;
                    }
                    result = this.context.SaveChanges() > 0;
                }

                return result;
            }
            catch (Exception ex)
            {
                this.log.Error(MethodBase.GetCurrentMethod().Name, ex);
                throw ex;
            }
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
