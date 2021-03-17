using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SharpLightReporting;

namespace SampleReporting
{
    public class InvoiceReportModel : IReportModel
    {
        public string InvoiceNumber
        {
            get
            {
                return "Test/001";
            }
        }

        public string InvoiceDate
        {
            get
            {
                string retval = DateTime.Now.Date.ToShortDateString();

                return retval;
            }
        }

        public string Consignor
        {
            get
            {
                string retval = "M/s Xyz Company Ltd";

                return retval;
            }
        }

        public string ConsignorAdd1
        {
            get
            {
                string retval = "Street ABC Block 1";

                return retval;
            }
        }

        public string ConsignorAdd2
        {
            get
            {
                string retval = "Landmark 123456";

                return retval;
            }
        }

        public string ConsignorAdd3
        {
            get
            {
                string retval = "";

                return retval;
            }
        }

        public string Gstin
        {
            get
            {
                string retval = "TX4356GSTIN987";

                return retval;
            }
        }

        public string NoOfPacks
        {
            get
            {
                string retval = "5";

                return retval;
            }
        }

        public string Transport
        {
            get
            {
                string retval = "FGH Roadways Ltd";

                return retval;
            }
        }

        public string LRNo
        {
            get
            {
                string retval = "FGH-1231";

                return retval;
            }
        }

        public string CurrentItemName
        {
            get
            {
                string retval = "Item 1 to 2";

                return retval;
            }
        }

        public string CurrentItemDescription
        {
            get
            {
                string retval = "Test Item Desc";

                return retval;
            }
        }

        public string CurrentItemHSN
        {
            get
            {
                string retval = "60002566";

                return retval;
            }
        }

        public string CurrentItemUnit
        {
            get
            {
                string retval = "Pcs";

                return retval;
            }
        }

        public decimal CurrentItemRate
        {
            get
            {
                decimal retval = 45.50m;

                return retval;
            }
        }

        public decimal CurrentItemQty
        {
            get
            {
                decimal retval = 2.0m;

                return retval;
            }
        }

        public decimal CurrentIncvoiceLineAmount
        {
            get
            {
                decimal retval = CurrentItemRate * CurrentItemQty;

                return retval;
            }
        }

        public decimal TotalAmount
        {
            get
            {
                decimal retval = CurrentIncvoiceLineAmount;

                return retval;
            }
        }
    }
}