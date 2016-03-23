using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Domain
{
    public abstract class LogBase
    {
        protected log4net.ILog log;

        public LogBase()
        {
            this.log = log4net.LogManager.GetLogger(this.GetType());
        }
    }
}
