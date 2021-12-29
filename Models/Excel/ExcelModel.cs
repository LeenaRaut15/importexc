using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using importexc.Data;

namespace importexc.Models.Excel
{
    public class ExcelModel
    {
        public int ID { get; set; }
        public string HospitalName { get; set; }
        public string Address { get; set; }
        public string AddmissionFee { get; set; }
    }
}