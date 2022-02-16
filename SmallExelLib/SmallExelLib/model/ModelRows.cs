using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallExelLib.model
{
    public class ModelRows
    {
        public ModelRows(DoubleValue height , int cell_num, string val, CellValues type, uint styleIndex)
        {
            this.height = height;
            this.cell_num = cell_num;
            this.val = val;
            this.type = type;
            this.styleIndex = styleIndex;
        }


        public DoubleValue height { get; set; }
        public int cell_num { get; set; }
        public string val { get; set; }
        public CellValues type { get; set; }
        public uint styleIndex { get; set; }
    }
}
