using Excel.Domain.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel.Domain.Entites;

namespace Excel.Domain.Concrete
{
    public class ValuationRepository : IValuationRepository
    {
        private EFDbContext context = new EFDbContext();

        public IEnumerable<Valuation> Valuations
        {
            get
            {
                return context.Valuations;
            }
        }
    }
}
