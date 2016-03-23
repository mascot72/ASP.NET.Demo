using Excel.Domain.Entites;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Domain.Concrete
{
    public class EFDbContext :DbContext
    {
        public DbSet<Valuation> Valuations { get; set; }

        public System.Data.Entity.DbSet<Excel.Domain.Entites.FileImport> FileImports { get; set; }

        public System.Data.Entity.DbSet<Excel.Domain.Entites.ExtendDefine> ExtendDefines { get; set; }

        public System.Data.Entity.DbSet<Excel.Domain.Entites.ExtendContent> ExtendContents { get; set; }

    }
}
