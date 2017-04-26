using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValidationTool
{
    class TableSetting
    {
        public string Name { get; set; }
        public string PrimaryKeyColumns { get; set; }
        public string OmittedColumnsMaster { get; set; }
        public string OmittedColumnsValidated { get; set; }
        public string SecondaryKeyColumns { get; set; }
        public int HeaderRow { get; set; }
        public bool ManualCalculation { get; set; }
    }
}
