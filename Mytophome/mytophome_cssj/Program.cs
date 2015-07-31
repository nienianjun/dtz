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
using System.Threading;
using HtmlAgilityPack;
using DTZ.Helper;

namespace mytophome_cssj
{
    class Program
    {
        //public static string strUrl = "http://so.gz.mytophome.com/search/searchProp.do?cityId=20&salesType=S&propertyName=&areaId=&subAreaId=&propertyType=Z&price=&isAsc=&key_word=&builtArea=&bedroom=&age=&lift=&circleId=&orderBy=&estateId=&num=-1&maxPageItems=20&pager.offset={0}";
        public static string strUrl = "http://gz.mytophome.com/prop/salelist/0_Z_0_0_0_0_0_0.html?&p={0}";

        static void Main(string[] args)
        {
            string strYearMonthDay, strYear, strMonth, strDay;

            Console.WriteLine("=========================出售数据分析器使用说明=========================");
            Console.WriteLine("选择查询模式");
            //Console.WriteLine("1：输入“1”按指定范围查询（此模式要求输入查询的开始日期，系统将自动查询从该日期起至今的所有出售数据）。");
            Console.WriteLine("1：输入“1”按指定日查询（此模式要求输入查询的指定日期，系统将自动查询该日期的所有出售数据）。");
            Console.WriteLine("2：直接回车查询当日的出售记录。");
            Console.WriteLine("");
            Console.WriteLine("注：输入的日期应该用中横线隔开，如2009-9-9");
            Console.WriteLine("========================================================================");
            Console.WriteLine("");
            Console.Write("请选择查询模式：");
            try
            {           
                string strType = Console.ReadLine();

                if (strType == "1")
                {
                    Console.Write("  请输入查询的指定日期：");
                    strYearMonthDay = Console.ReadLine();
                    try
                    {
                        strYear = strYearMonthDay.Split('-')[0];
                        strMonth = strYearMonthDay.Split('-')[1];
                        strDay = strYearMonthDay.Split('-')[2];
                    }
                    catch
                    {
                        Console.Write("错误，您输入的日期格式不正确。");
                        Console.Read();
                        return;
                    }
                    SaveFileXls(strYear, strMonth.ToString(), strDay.ToString());
                }
                else
                {
                    SaveFileXls(DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString());
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("程序出错：" + ex.ToString());
                Console.ReadKey(); ;
            }

            Console.ReadKey(); ;
        }

        static void SaveFileXls(string strYear, string strMonth, string strDay)
        {
            Console.WriteLine("");
            Console.WriteLine("开始查询" + strYear + "年" + strMonth + "月" + strDay + "日的出售数据");

            List<EstateSell> listEstateSell = new List<EstateSell>();

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

            wbook.CreateSheet("今日出售数据");
            wsheet = wbook.GetSheet("今日出售数据");
            wbook.SetActiveSheet = "今日出售数据";
            style = wbook.CreateStyle();
            style.Pattern = EnumFill.Solid;
            style.PatternForeColour = EnumColours.Grey25;
            style.Font.Size = 11;
            style.Font.Bold = true;

            string[] strDBBP_Title = { "标题", "区域", "楼盘", "户型", "面积(㎡)", "成交价(万)", "单价(元/㎡)", "时间" };
            for (int o = 0; o < strDBBP_Title.Length; o++)
            {
                wsheet.AddCell((ushort)(o + 1), 1, strDBBP_Title[o], style);
            }

            int intRow = 1;
            int intMaxPage = 1000;
            int intPage = 0;
            int intOffset = 0;
            int intError = 0;

            bool blState = true;

            do
            {
                intPage++;
                intOffset = (intPage - 1) * 24;
                //intOffset = 1080;

                try
                {
                    //if (intPage == 1)
                    //{
                    //    string strMaxPage = GetNoHtml(GetstrCenter(strHtml, "共<span class=\"zi_333333_12\">", "</span>"));
                    //    intMaxPage = int.Parse(Math.Ceiling(int.Parse(strMaxPage) / 24.0).ToString());
                    //}
                    //Console.WriteLine("intMaxPage : " + intMaxPage);
                    string s = GetXmlHttp(string.Format(strUrl, intOffset.ToString()));
                    List<EstateSell> listTemp = GetarrEstateSell(s);      

                    DateTime inputSJ = DateTime.Parse(strYear + "-" + strMonth + "-" + strDay);
                    for (int o = 0; o < listTemp.Count; o++)
                    {
                        if (listTemp[o] != null)
                        {
                            DateTime searchSJ = DateTime.Parse(listTemp[o].SJ);
                            if (inputSJ == searchSJ)
                            {
                                listEstateSell.Add(listTemp[o]);
                            }
                            else if (searchSJ < inputSJ)
                            {
                                blState = false;
                                break;
                            }
                        }
                        else
                            break;
                    }

                    intError = 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("");
                    Console.WriteLine("" + "第" + intPage + "页数据分析出错。错误提示：" + ex.ToString());

                    intError++;
                    if (intError < 1)
                    {
                        Console.WriteLine("等待20秒后继续运行...");
                        Thread.Sleep(20000);
                        Console.WriteLine("");
                        intPage--;
                    }
                }

            }
            while (intPage < intMaxPage & blState);

            foreach (EstateSell es in listEstateSell)
            {
                intRow++;
                ushort intX = (ushort)intRow;
                wsheet.AddCell(1, intX, es.BT);
                wsheet.AddCell(2, intX, es.QY);
                wsheet.AddCell(3, intX, es.LP);
                wsheet.AddCell(4, intX, es.HX);
                wsheet.AddCell(5, intX, es.MJ);
                wsheet.AddCell(6, intX, es.CJJ);
                wsheet.AddCell(7, intX, es.DJ);
                wsheet.AddCell(8, intX, es.SJ);
            }

            string strPath = @"Xls\";
            if (!Directory.Exists(strPath))
            {
                Directory.CreateDirectory(strPath);
            }
            strPath = strPath + "出售数据" + strYear + "-" + strMonth + "-" + strDay + ".xls"; //保存的路径和文件名

            Console.WriteLine("");
            Console.WriteLine("" + strYear + "年" + strMonth + "月" + strDay + "日的数据已经产生。");

            wbook.Save(strPath);
            watch.Stop();
        }

        static List<EstateSell> GetarrEstateSell(String pageHtml)
        {
            List<EstateSell> listES = new List<EstateSell>();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(pageHtml);
            HtmlNodeCollection nodes = doc.GetElementbyId("style1").SelectNodes("//li[@class='lb_list01']");
            foreach (HtmlNode nodeTemp in nodes)            
            {
                HtmlDocument htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(nodeTemp.InnerHtml);
                HtmlNode node = htmlDoc.DocumentNode;

                EstateSell estateSell = new EstateSell();

                estateSell.BT = node.SelectSingleNode("//li[@class='two_word_li01'][1]/a[1]").InnerText;
                //Console.WriteLine("标题：" + estateSell.BT);

                String sTemp = node.SelectSingleNode("//div[@class='fy_di01'][1]").InnerText.Replace("&nbsp;", " ");
                //sTemp = StrHelper.ClearHTML(sTemp);
                estateSell.LP = sTemp.Split('[')[0].Trim();
                estateSell.QY = sTemp.Split('[')[1].Split(' ')[0];
                //Console.WriteLine("区域：" + estateSell.QY);
                //Console.WriteLine("楼盘：" + estateSell.LP);

                estateSell.HX = node.SelectSingleNode("//div[@class='two_li02_div02'][1]").InnerText;
                //Console.WriteLine("户型：" + estateSell.HX);

                estateSell.MJ = node.SelectSingleNode("//div[@class='two_li02_div02'][2]").InnerText.Replace("平方", "");
                //Console.WriteLine("面积：" + estateSell.MJ);

                estateSell.CJJ = node.SelectSingleNode("//div[@class='two_li02_div03_w'][1]/span[1]").InnerText;
                //Console.WriteLine("成交价：" + estateSell.CJJ);
                estateSell.DJ = node.SelectSingleNode("//div[@class='two_li02_div03_w'][2]").InnerText.Replace("元/平方", "");
                //Console.WriteLine("单价：" + estateSell.DJ);

                sTemp = node.SelectSingleNode("//span[@class='a03'][1]").InnerText.Trim();
                if (sTemp.IndexOf("天") >= 0)
                {
                    sTemp = Regex.Split(sTemp, "天", RegexOptions.IgnoreCase)[0].Substring(1);
                    estateSell.SJ = DateTime.Now.AddDays(-int.Parse(sTemp)).ToString("yyyy-MM-dd");
                }
                else if (sTemp.IndexOf("小时") >= 0)
                {
                    sTemp = Regex.Split(sTemp, "小时", RegexOptions.IgnoreCase)[0].Substring(1);
                    estateSell.SJ = DateTime.Now.AddHours(-int.Parse(sTemp)).ToString("yyyy-MM-dd");
                }
                else if (sTemp.IndexOf("分钟") >= 0)
                {
                    sTemp = Regex.Split(sTemp, "分钟", RegexOptions.IgnoreCase)[0].Substring(1);
                    estateSell.SJ = DateTime.Now.AddMinutes(-int.Parse(sTemp)).ToString("yyyy-MM-dd");
                }
                else
                {
                    estateSell.SJ = DateTime.Now.ToString("yyyy-MM-dd");
                }
                //Console.WriteLine("时间：" + estateSell.SJ);
                listES.Add(estateSell);
                node.Clone();
                htmlDoc = null;
            }
            nodes.Clear();
            doc = null;
            return listES;
        }


        /// <summary>
        /// 读取指定Url的源代码
        /// </summary>
        /// <param name="strUrl"></param>
        /// <returns></returns>
        static string GetXmlHttp(string strUrl)
        {
            return UrlHelper.GetUrlHtml(strUrl, "GBK");            
            //return FileHelper.ReadFile("temp.html","UTF-8");
        }

    }
}
