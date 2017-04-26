using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValidationTool
{
    class SQLToSQLSingleSingleConfig
    {
        public string FirstQueryConnectionString { get; set; }
        public string FirstQuery { get; set; }
        public string SecondQueryConnectionString { get; set; }
        public string SecondQuery { get; set; }
        public string KeyColumns { get; set; }
        public string SecondaryKeyColumns { get; set; }
    }
}
