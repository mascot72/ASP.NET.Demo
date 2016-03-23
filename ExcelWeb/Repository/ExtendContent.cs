using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelWeb.Repository
{
    public class ExtendContent
    {
        public int ImportID { get; set; }

        //VALU_EXCEL_EXT(ID)        
        public int EID { get; set; }
        
        public string Content { get; set; }
        
        public int Ref1 { get; set; }
        
        public string Ref2 { get; set; }

        public virtual Valuation Valuation { get; set; }

        public virtual ExtendDefine ExtendDefine { get; set; }
        
    }
}