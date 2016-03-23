using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Excel.Domain.Entites
{
    public class ExtendDefine
    {
        //Primary Key column
        [Key]
        public int ID { get; set; }
        
        public string Name { get; set; }
        
        public DateTime? CreateDate { get; set; }

    }

}