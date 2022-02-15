using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmallExelLib.model
{
    public class ModelColumn
    {
        public ModelColumn(UInt32Value Min, UInt32Value Max, DoubleValue Width, BooleanValue CustomWidth)
        {
            this.Min = Min;
            this.Max = Max;
            this.Width = Width;
            this.CustomWidth = CustomWidth;
        }
        public UInt32Value Min { get; set; }
        public UInt32Value Max { get; set; }

        public DoubleValue Width { get; set; }

        public BooleanValue CustomWidth { get; set; }
    }

}
