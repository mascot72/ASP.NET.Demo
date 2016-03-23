﻿using Excel.Domain.Entites;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Domain.Abstract
{
    interface IExtendContentRepository
    {
        IEnumerable<ExtendContent> extendContent { get; }
    }
}
