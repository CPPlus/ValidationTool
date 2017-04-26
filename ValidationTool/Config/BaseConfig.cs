using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValidationTool
{
    class BaseConfig
    {
        public string ValidationToolConnectionString { get; set; }
        public string ComparisonSourceTypes { get; set; }
        public bool SmartRounding { get; set; }
        public bool SmartRoundingFormatting { get; set; }
        public decimal SmartRoundingDelta { get; set; }
        public string[] ValuesToNeglect { get; set; }
        public string[] SubstringsToNeglect { get; set; }
    }
}
