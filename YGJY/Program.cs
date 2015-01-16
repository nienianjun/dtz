using System;
using System.Collections.Generic;
using System.Text;
using DTZ.Helper;
using DTZ.YGJY.Models;
using HtmlAgilityPack;

namespace DTZ.YGJY
{
    class Program
    {
        private static string urlReferer = "http://g4c.laho.gov.cn/";
        private static string urlEncoding = "gb2312";

        private static string logFileName = "historycs.log";

        static void Main(string[] args)
        {
            Console.Write("程序启动");
            Console.WriteLine("");
            try
            {
                UrlHelper.SetEncoding(urlEncoding);
                UrlHelper.SetReferer(urlReferer);
                //test();
                HistorycsLog historycsLog = new HistorycsLog();
                if (FileHelper.FileExists(logFileName))
                {
                    historycsLog = StrHelper.FromJson<HistorycsLog>(FileHelper.ReadFile(logFileName));
                    if (historycsLog.ProjectID.Count > 200)
                    {
                        historycsLog.ProjectID.RemoveRange(0, historycsLog.ProjectID.Count - 200);
                    }
                }
                Services.ProjectServices PS = new Services.ProjectServices();

                List<Project> listProject = new List<Project>();

                Console.Write("请输入要导出的者预售证年月(如2012-09)或者预售证号 ：");
                string presellNo = Console.ReadLine();
                while (presellNo == "")
                {
                    Console.Write("   请输入要导出的预售证查询条件：");
                    presellNo = Console.ReadLine();
                }
                
                if (presellNo.IndexOf("-") != -1)
                {   //指定月份预售证年月
                    string awardedStartDate = presellNo + "-01";
                    string awardedEndDate = DateTime.Parse(awardedStartDate).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
                    Console.WriteLine(" 下面开始导入预售发证时间在" + awardedStartDate + "至" + awardedEndDate + "之间的项目。");
                    //Console.Read();
                    //Console.WriteLine(" 上次自动运行时间：" + historycsLog.RunDate);

                    Console.WriteLine(" 项目加载中......");
                    listProject = PS.getProjectList(awardedStartDate, awardedEndDate);
                    Console.WriteLine(" 共加载到" + listProject.Count + "个项目。");
                    //foreach (Project objPS in PS.getProjectList())
                    //{
                    //    bool projectExists = false;
                    //    foreach (string id in historycsLog.ProjectID)
                    //    {
                    //        if (objPS.ProjectID == id)
                    //        {
                    //            projectExists = true;
                    //            break;
                    //        }
                    //    }
                    //    if (!projectExists)
                    //    {
                    //        listProject.Add(objPS);
                    //    }
                    //}
                }
                else
                {
                    Console.WriteLine(" 下面开始导入预售证号为：" + presellNo + "的项目。");
                    //Console.Read();
                    listProject = PS.getProjectList(presellNo);
                }

                foreach (Project pro in listProject)
                {
                    try
                    {
                        Console.WriteLine("");
                        Console.WriteLine(" 【" + pro.ProjectName + "】开始运行");
                        PS.SaveExcel(PS.supplementProject(pro));
                        Console.WriteLine(" 【" + pro.ProjectName + "】已经完成");

                        historycsLog.ProjectID.Add(pro.ProjectID);
                        historycsLog.RunDate = DateTime.Now.ToString("yyyy-MM-dd");
                        FileHelper.SaveFile(StrHelper.ToJson(historycsLog), logFileName);

                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(" 【" + pro.ProjectName + "】运行失败");
                        Console.WriteLine("    失败原因：" + e.ToString());
                        System.Threading.Thread.Sleep(5000);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("    程序执行出错：" + e.ToString());
                Console.ReadLine();
            }
            Console.WriteLine("");
            Console.Write("程序结束");
            System.Threading.Thread.Sleep(5000);
        }

        static void test()
        {
            Console.WriteLine("test start");
            //string html = UrlHelper.GetUrlHtml("http://g4c.laho.gov.cn/search/project/sellFormDetail.jsp?unitID=100001582122");
            //FileHelper.SaveFile(html, "D:\\logs\\temp10.html");

            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(FileHelper.ReadFile("D:\\logs\\100000016697.html"));
            //Console.WriteLine(doc.DocumentNode.SelectNodes("//table[1]/tr[position() > 2]")[0].InnerHtml);
            HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("/html[1]/body[1]//div[1]/table[1]/tr");
            Console.WriteLine(nodes.Count);

            //Console.ReadLine();
            int i = 0;
            //HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//table[2]/tr[position() > 2]/td[@class=\"tab_style01_td\"]");
            foreach (HtmlNode node in nodes)
            {
                Console.WriteLine( i.ToString() + " : " + node.InnerHtml.Trim());
                i++;
                Console.ReadLine();
            }
            Console.WriteLine("test end");
            Console.ReadLine();

        }
    }
}
