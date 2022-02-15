using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallExelLib.model
{
    public class DataModelSheet
    {
        public DataModelSheet(List<ModelColumn> columnList, List<ModelHeaderColumn> headerColumnList, List<List<ModelRows>> rowsColumnList)
        {
            this.columnList = columnList;
            this.headerColumnList = headerColumnList;
            this.rowsColumnList = rowsColumnList;

        }

        public List<ModelColumn> columnList { get; set; }
        public List<ModelHeaderColumn> headerColumnList { get; set; }
        public List<List<ModelRows>> rowsColumnList { get; set; }
    }
}
