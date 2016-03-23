using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace Excel.Domain.Entites
{
    public class ExtendContent
    {
        [Key]
        public int ID { get; set; }

        [ForeignKey("Valuation")]
        [Column(Order = 1)]
        public int ImportID { get; set; }

        //VALU_EXCEL_EXT(ID)
        [ForeignKey("ExtendDefine")]
        [Column(Order = 2)]
        public int EID { get; set; }
        
        public string Content { get; set; }
        
        public int Ref1 { get; set; }
        
        public string Ref2 { get; set; }

        public virtual Valuation Valuation { get; set; }

        public virtual ExtendDefine ExtendDefine { get; set; }
        
    }
}