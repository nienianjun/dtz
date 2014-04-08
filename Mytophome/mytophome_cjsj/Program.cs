using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Net;
using System.Xml;
using System.Data;
using System.Diagnostics;
using Biff8Excel;
using Biff8Excel.Excel;
using System.Web;
using DataModel;

namespace mytophome_cjsj
{
    class Program
    {
        public static string strUrl = "http://gz.mytophome.com/deal/queryByMonth.do?cityId=20&month={1}&year={0}&jidu=0&halfYear=0&estateName=&areaId=&startDate=&endDate=&maxPageItems=20&pager.offset={2}";

        static void Main(string[] args)
        {
            string strYearMonth, strYear, strMonth;

            Console.WriteLine("=========================成交数据分析器使用说明=========================");
            Console.WriteLine("选择查询模式");
            Console.WriteLine("1：输入“1”按指定范围查询（此模式要求输入查询的开始月份，系统将自动查询从该月起至今所有成交记录）。");
            Console.WriteLine("2：输入“2”按指定月份查询（此模式要求输入指定月份，系统将自动查询该月所有成交记录）。");
            Console.WriteLine("3：直接回车查询上个月成交记录。");
            Console.WriteLine("");
            Console.WriteLine("注：输入的年月应该用中横线隔开，如2009-9");
            Console.WriteLine("========================================================================");
            Console.WriteLine("");
            Console.Write("请选择查询模式：");

            try
            {
                string strType = Console.ReadLine();
                if (strType == "1")
                {
                    Console.Write("  请输入查询的开始年月：");
                    strYearMonth = Console.ReadLine();
                    if (strYearMonth.IndexOf('-') == -1)
                    {
                        Console.Write("错误，您输入的年月格式不正确。");
                        Console.Read();
                        return;
                    }

                    strYear = strYearMonth.Split('-')[0];
                    strMonth = strYearMonth.Split('-')[1];
                    for (int iYear = int.Parse(strYear); iYear <= DateTime.Now.Year; iYear++)
                    {
                        for (int iMonth = int.Parse(strMonth); iMonth <= 12; iMonth++)
                        {
                            SaveFileXls(iYear.ToString(), iMonth.ToString());
                            if (iYear == DateTime.Now.Year && iMonth == DateTime.Now.Month)
                            {
                                return;
                            }
                        }
                    }
                }
                else if (strType == "2")
                {
                    Console.Write("  请输入查询的指定年月：");
                    strYearMonth = Console.ReadLine();
                    if (strYearMonth.IndexOf('-') == -1)
                    {
                        Console.Write("错误，您输入的年月格式不正确。");
                        return;
                    }

                    strYear = strYearMonth.Split('-')[0];
                    strMonth = strYearMonth.Split('-')[1];
                    SaveFileXls(strYear, strMonth);

                }
                else
                    SaveFileXls(DateTime.Now.Year.ToString(), DateTime.Now.AddMonths(-1).Month.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("程序出错：" + ex.ToString());
                Console.ReadKey(); ;
            }
        }
        static void SaveFileXls(string strYear, string strMonth)
        {
            Console.WriteLine("");
            Console.WriteLine("开始查询" + strYear + "年" + strMonth + "月数据");

            string strHtml = "";
            Estate[] arrEstate;

            Stopwatch watch = new Stopwatch();
            watch.Start();
            ExcelWorkbook wbook;
            ExcelWorksheet wsheet;
            ExcelCellStyle style;
            //style.BottomLineStyle = EnumLineStyle.Thick;
            //style.TopLineStyle = EnumLineStyle.Medium;
            DateTime dt = System.DateTime.Now;
            wbook = new ExcelWorkbook();
            wbook.SetDefaultFont("Arial", 10);

            wbook.CreateSheet("详细成交数据");
            wsheet = wbook.GetSheet("详细成交数据");
            wbook.SetActiveSheet = "详细成交数据";
            style = wbook.CreateStyle();
            style.Pattern = EnumFill.Solid;
            style.PatternForeColour = EnumColours.Grey25;
            style.Font.Size = 11;
            style.Font.Bold = true;

            string[] strDBBP_Title = { "区域", "物业地址", "栋苑", "面积(㎡)", "成交价(万)", "单价(元/㎡)", "成交日期" };
            for (int o = 0; o < strDBBP_Title.Length; o++)
            {
                wsheet.AddCell((ushort)(o + 1), 1, strDBBP_Title[o], style);
            }

            int intRow = 1;
            int intMaxPage = 5;
            int intPage = 0;
            int intOffset = 0;

            do
            {
                intPage++;
                intOffset = (intPage - 1) * 20;
                strHtml = GetXmlHttp(string.Format(strUrl, strYear, strMonth, intOffset.ToString()));

                arrEstate = GetEstate(strHtml);
                for (int o = 0; o < arrEstate.Length; o++)
                {
                    if (arrEstate[o] != null)
                    {
                        intRow++;
                        ushort intX = (ushort)intRow;
                        wsheet.AddCell(1, intX, arrEstate[o].QY);
                        wsheet.AddCell(2, intX, arrEstate[o].WYDZ);
                        wsheet.AddCell(3, intX, arrEstate[o].DY);
                        wsheet.AddCell(4, intX, arrEstate[o].MJ);
                        wsheet.AddCell(5, intX, arrEstate[o].CJJ);
                        wsheet.AddCell(6, intX, arrEstate[o].DJ);
                        wsheet.AddCell(7, intX, arrEstate[o].CJRQ);
                    }
                    else
                        break;
                }

                if (intPage == 1)
                {
                    string strMaxPage = GetNoHtml(GetstrCenter(strHtml, "pageJump_Start", "pageJump_End"));
                    strMaxPage = GetstrCenter(strMaxPage, "个主题 第", "页");
                    intMaxPage = int.Parse(strMaxPage.Split('/')[1]);
                }
            }
            while (intPage < intMaxPage);

            string strPath = @"Xls\";
            if (!Directory.Exists(strPath))
            {
                Directory.CreateDirectory(strPath);
            }
            strPath = strPath + "成交数据" + strYear + "_" + strMonth + ".xls"; //保存的路径和文件名

            Console.WriteLine("" + strYear + "年" + strMonth + "月数据已经产生。");

            wbook.Save(strPath);
            watch.Stop();
        }

        static Estate[] GetEstate(string strHtml)
        {
            Estate[] arrEstate_Temp = new Estate[100];

            strHtml = GetstrCenter(strHtml, "class=\"deD_ctt\"", "</ul>");
            int iLength = Regex.Split(strHtml, "<li>", RegexOptions.IgnoreCase).Length - 1;
            string strTemp = "";
            for (int i = 0; i < iLength; i++)
            {
                strTemp = Regex.Split(strHtml, "<li>", RegexOptions.IgnoreCase)[i + 1];
                arrEstate_Temp[i] = new Estate();
                arrEstate_Temp[i].QY = GetstrCenter(strTemp, "class=\"deD_houseArea\">", "</span>");
                arrEstate_Temp[i].WYDZ = GetNoHtml(GetstrCenter(strTemp, "class=\"deD_add\">", "</span>"));
                arrEstate_Temp[i].DY = GetstrCenter(strTemp, "class=\"deD_building\">", "</span>");
                arrEstate_Temp[i].MJ = GetstrCenter(strTemp, "class=\"deD_houseMagnitude\">", "</span>");
                arrEstate_Temp[i].CJJ = GetstrCenter(strTemp, "class=\"deD_closePrice\">", "</span>");
                arrEstate_Temp[i].DJ = GetstrCenter(strTemp, "class=\"deD_unitPrice\">", "</span>");
                arrEstate_Temp[i].CJRQ = GetstrCenter(strTemp, "class=\"deD_closeDate\">", "</span>").Substring(0, 10);

                //Console.WriteLine("QY:" + arrEstate_Temp[i].QY);
                //Console.WriteLine("WYDZ:" + arrEstate_Temp[i].WYDZ);
                //Console.WriteLine("DY:" + arrEstate_Temp[i].DY);
                //Console.WriteLine("MJ:" + arrEstate_Temp[i].MJ);
                //Console.WriteLine("CJJ:" + arrEstate_Temp[i].CJJ);
                //Console.WriteLine("DJ:" + arrEstate_Temp[i].DJ);
                //Console.WriteLine("CJRQ:" + arrEstate_Temp[i].CJRQ);
            }
            return arrEstate_Temp;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strTest"></param>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        static string GetstrCenter(string strTest, string str1, string str2)
        {
            string strTemp = "";
            strTemp = Regex.Split(strTest, str1, RegexOptions.IgnoreCase)[1];
            strTemp = Regex.Split(strTemp, str2, RegexOptions.IgnoreCase)[0];

            return strTemp;
        }

        static string GetNoHtml(string strHtml)
        {
            return Regex.Replace(strHtml, "<.+?>", "", RegexOptions.Singleline);
        }


        /// <summary>
        /// 读取指定Url的源代码
        /// </summary>
        /// <param name="strUrl"></param>
        /// <returns></returns>
        static string GetXmlHttp(string strUrl)
        {
            try
            {
                Console.WriteLine(strUrl);
                Uri userUri = new Uri(strUrl);
                WebRequest userRequest = WebRequest.Create(userUri);
                WebResponse userResponse = userRequest.GetResponse();
                Encoding myEncoding = System.Text.Encoding.GetEncoding(0);
                StreamReader readStream = new StreamReader(userResponse.GetResponseStream(), myEncoding);
                String strHTML = readStream.ReadToEnd();
                readStream.Close();
                return strHTML;

            }
            catch (OverflowException e)
            {
                Console.WriteLine("{0}", e.Message);
                Console.ReadLine();
                return "";
            }
        }

    }
}
