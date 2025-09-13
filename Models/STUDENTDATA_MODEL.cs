using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCPRACTICES.Models
{
    public class STUDENTDATA_MODEL
    {
        public int sr_no { get; set; }
        public int id { get; set; }
        public string studentname { get; set; }
        public int studentage { get; set; }
        public string studentqualification { get; set; }
        public string studentgender { get; set; }
        public int countryid { get; set; }
        public int stateid { get; set; }
        public string countryname { get; set; }
        public string statename { get; set; }

    }
}