using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LandedCosts
{
    public class LandedCostsModel
    {
        public string VendorCode { get; set; }
        public int InvoiceDocEntry { get; set; }
        public DateTime PostingDate { get; set; }
        public string Number { get; set; }
        public string Comment { get; set; }

        public List<LandedCostsRowModel> Rows = new List<LandedCostsRowModel>();
    }

    public class LandedCostsRowModel
    {
        public string LandedCostCode { get; set; }
        public string LandedCostName { get; set; }
        public double Price { get; set; }
        public string Currency { get; set; }
        public double Rate { get; set; }
        public string VatGroup { get; set; }
        public string Broker { get; set; }
    }

}
