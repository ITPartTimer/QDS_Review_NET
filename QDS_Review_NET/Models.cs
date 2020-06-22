using System;
using System.Collections.Generic;
using System.Text;

namespace QDS_Review_NET.Models
{
    public class EmployeesReportsModel
    {
        public int rptID { get; set; }
        public string email { get; set; }
        public string temppath { get; set; }
        public string rootpath { get; set; }
        public string filename { get; set; }
        public string fullpath { get; set; }
    }

    public class QDSDataModel
    {
        public string whs { get; set; }
        public string tag { get; set; } 
        public string typ { get; set; }
        public string sts { get; set; }
        public string ref_pfx { get; set; }
        public string ref_no { get; set; }
        public string qds { get; set; }
        public string dt { get; set; }
    }
}
