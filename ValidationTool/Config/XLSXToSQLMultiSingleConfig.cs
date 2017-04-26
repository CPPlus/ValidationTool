using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValidationTool
{
    class XLSXToSQLMultiSingleConfig
    {
        public string MasterFilePath { get; set; }
        public string ConnectionString { get; set; }
        public string Schema { get; set; }
        public string Prefix { get; set; }
        public SubstringMap[] SubstringMaps { get; set; }
        public string Suffix { get; set; }
    }
}
