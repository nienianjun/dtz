using System;
using System.Collections.Generic;
using System.Text;

namespace DTZ.YGJY.Models
{
    public class Project
    {
        public string ProjectID { get; set; }
        public string ProjectPresell { get; set; }
        public string ProjectName { get; set; }
        public string ProjectUrl { get; set; }

        public Project()
        {

        }

        public DB_ESTATE Estate { get; set; }

        public List<DB_FLAT> FlatList { get; set; }

        public List<DB_BLOCK> BlockList { get; set; }
    }
}
