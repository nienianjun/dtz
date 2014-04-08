using System;
using System.Collections.Generic;
using System.Text;

namespace DTZ.YGJY.Models
{
    public class DB_ESTATE
    {
        /// <summary>
        /// 物业ID[1]
        /// </summary>
        public string WYID { get; set; }

        /// <summary>
        /// 城市[2]
        /// </summary>
        public string CS { get; set; }

        /// <summary>
        /// 行政区[3] 片区[5]
        /// </summary>
        public string XZQ { get; set; }

        /// <summary>
        /// 物业名称[7]
        /// </summary>
        public string WYMC { get; set; }

        /// <summary>
        /// 物业地址[11]
        /// </summary>
        public string WYDZ { get; set; }

        /// <summary>
        /// 占地面积[16]
        /// </summary>
        public string ZDMJ { get; set; }

        /// <summary>
        /// 总建筑面积[17]
        /// </summary>
        public string ZJZMJ { get; set; }



    }
}
