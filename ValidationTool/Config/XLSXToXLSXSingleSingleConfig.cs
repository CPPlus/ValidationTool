using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValidationTool
{
    class XLSXToXLSXSingleSingleConfig
    {
        public string FirstXLSXFilePath { get; set; }
        public string FirstXLSXFileSheet { get; set; }
        public string SecondXLSXFilePath { get; set; }
        public string SecondXLSXFileSheet { get; set; }
        public string KeyColumns { get; set; }
        public string SecondaryKeyColumns { get; set; }
        public bool ManualCalculation { get; set; }
    }
}
