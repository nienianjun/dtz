using System;
using System.Collections.Generic;
using System.Text;

namespace DTZ.YGJY.Models
{
    public class DB_FLAT
    {

        /// <summary>
        /// 单元ID
        /// </summary>
        public string _unitID { get; set; }

        /// <summary>
        /// 大楼ID[1]
        /// </summary>
        public string DLID { get; set; }
        /// <summary>
        /// 层数列(数字)[2]
        /// </summary>
        public string CSL { get; set; }
        /// <summary>
        /// 层数名[3]
        /// </summary>
        public string CSM { get; set; }
        /// <summary>
        /// 单元[5]
        /// </summary>
        public string DY { get; set; }
        /// <summary>
        /// 建筑面积(平方米)[6]
        /// </summary>
        public string JZMJ { get; set; }
        /// <summary>
        /// 套内面积(平方米)[7]
        /// </summary>
        public string TNMJ { get; set; }

        /// <summary>
        /// 用途(L1)[18]
        /// </summary>
        public string YTL1 { get; set; }

        /// <summary>
        /// 房[20]
        /// </summary>
        public string F { get; set; }
        /// <summary>
        /// 厅[21]
        /// </summary>
        public string T { get; set; }
    }
}
