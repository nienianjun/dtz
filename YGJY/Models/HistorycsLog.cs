using System;
using System.Collections.Generic;
using System.Text;

namespace DTZ.YGJY.Models
{
    public class HistorycsLog
    {
        public string RunDate { get; set; }
        public List<String> ProjectID { get; set; }

        public HistorycsLog()
        {
            RunDate = DateTime.Now.ToString("yyyy-MM-dd");
            ProjectID = new List<string>();
        }
    }
}
