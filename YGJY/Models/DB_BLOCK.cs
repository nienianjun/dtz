using System;
using System.Collections.Generic;
using System.Text;

namespace DTZ.YGJY.Models
{
    public class DB_BLOCK
    {
        public string _BuildingID { get; set; }

        /// <summary>
        /// 物业ID[1]
        /// </summary>
        public string WYID { get; set; }

        /// <summary>
        /// 大楼ID[2]
        /// </summary>
        public string DLID { get; set; }

        /// <summary>
        /// 楼栋中文名称[5]
        /// </summary>
        public string LDZWMC { get; set; }

        /// <summary>
        /// 竣工日期[11]
        /// </summary>
        public string JGRQ { get; set; }

        /// <summary>
        /// 发展商[25]
        /// </summary>
        public string FZS { get; set; }

        ///// <summary>
        ///// 单栋大楼名称地址[26]
        ///// </summary>
        //public string DDDLMCDZ { get; set; }
    }
}
