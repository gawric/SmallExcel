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
        public ModelRows(int cell_num, string val, CellValues type, uint styleIndex)
        {
           // this.RowIndex = RowIndex;
            this.cell_num = cell_num;
            this.val = val;
            this.type = type;
            this.styleIndex = styleIndex;
        }

   

        public int cell_num { get; set; }
        public string val { get; set; }
        public CellValues type { get; set; }
        public uint styleIndex { get; set; }
    }
}
