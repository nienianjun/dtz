using System;
using System.Collections.Generic;
using System.Text;

namespace DTZ.YGJY.Models
{
    public class DB_BUILDING_PROFILE
    {
        /// <summary>
        /// 物业ID[0]
        /// </summary>
        public String EstateID { get; set; }

        /// <summary>
        /// 大楼ID{1}
        /// </summary>
        public string DLID { get; set; }

        /// <summary>
        /// 单栋大楼中文名称
        /// </summary>
        public string DDDLZWMC { get; set; }
        /// <summary>
        /// 竣工日期
        /// </summary>
        public string JGRQ { get; set; }
        /// <summary>
        /// 用途(L1)
        /// </summary>
        public string YTL1 { get; set; }
        /// <summary>
        /// 用途(L2)
        /// </summary>
        public string YTL2 { get; set; }
        /// <summary>
        /// 发展商
        /// </summary>
        public string FZS { get; set; }
        /// <summary>
        /// 单栋大楼名称地址
        /// </summary>
        public string DDDLMCDZ { get; set; }
    }
}
