﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Excel.Domain.Entites
{
    public class FileImport
    {
        //Primary Key column
        [Key]
        public int ID { get; set; }

        public string Path { get; set; }

        public string Name { get; set; }

        public string ExtName { get; set; }

        public string Result { get; set; }

        public string Reason { get; set; }

        public string Remark { get; set; }

        public string Extend { get; set; }

        public DateTime? CreateDate { get; set; }

        public DateTime? UpdateDate { get; set; }

        public string Creator { get; set; }

        public string Updater { get; set; }

        public double Size { get; set; }

        public ICollection<Valuation> Valuation { get; set; }

    }
}