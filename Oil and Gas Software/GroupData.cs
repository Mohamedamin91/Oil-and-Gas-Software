using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Oil_and_Gas_Software
{
   public class GroupData
    {
        public int Material { get; set; }
        public int Qty { get; set; }
        public int Pqty { get; set; }
        public int AmountSum { get; set; }
        public IGrouping<decimal, System.Data.DataRow> Data { get; set; }
    }
}
