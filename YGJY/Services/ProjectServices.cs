using System;
using System.Collections.Generic;
using System.Text;
using DTZ.YGJY.Models;
using DTZ.Helper;
using HtmlAgilityPack;
using Biff8Excel.Excel;
using System.Diagnostics;
using Biff8Excel;
using System.IO;

namespace DTZ.YGJY.Services
{ 
    public class ProjectServices
    {
        private string ProjectDomain = "http://g4c.laho.gov.cn";

        public List<Project> getProjectList()
        {
            List<Project> listProject = new List<Project>();

            string html = UrlHelper.GetUrlHtml(ProjectDomain + "/");
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//div[@id=\"left\"]//table[@class=\"box_tab_style01 lh30 mt10\"][position() <= 2]/tr[position()>1]");
            foreach (HtmlNode node in nodes)
            {
                Project pro = new Project();
                HtmlNode proNode = node.SelectSingleNode("td/a");
                pro.ProjectID = StrHelper.GetRegexValue(proNode.Attributes["href"].Value, "&pjID=(?<1>\\d{1,20})&")[1];
                pro.ProjectPresell = node.SelectSingleNode("td[2]/a").InnerText;
                pro.ProjectName = proNode.Attributes["title"].Value;
                pro.ProjectUrl = ProjectDomain + proNode.Attributes["href"].Value;
                listProject.Add(pro);
            }
            return listProject;
        }
        public List<Project> getProjectList(String awardedStartDate, String awardedEndDate)
        {
            List<Project> listProject = new List<Project>();
            //string html = UrlHelper.GetUrlHtml(ProjectDomain + "/search/presell/preSellSearch.jsp?awardedStartDate=" + awardedStartDate + "&awardedEndDate=" + awardedEndDate + "&judge=1&currPage=0");
            //string html = FileHelper.ReadFile("D:\\logs\\temp1.html");

            int pageCount = 1;
            for (int iPage = 0; iPage < pageCount; iPage++)
            {
                string html = UrlHelper.GetUrlHtml(ProjectDomain + "/search/presell/preSellSearch.jsp?awardedStartDate=" + awardedStartDate + "&awardedEndDate=" + awardedEndDate + "&judge=1&currPage=" + iPage.ToString());
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(html);
                pageCount = int.Parse(StrHelper.GetRegexValue(doc.DocumentNode.SelectSingleNode("//div[@id=\"main\"]/div[2]/table[5]/tbody/tr[1]/td[1]").InnerText, "总共(?<1>\\d{1,3})页")[1]);
                //html = FileHelper.ReadFile("D:\\logs\\temp1.html");
                HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//div[@id=\"main\"]/div[2]/table[4]/tr");
                foreach (HtmlNode node in nodes)
                {
                    HtmlDocument hd = new HtmlDocument();
                    hd.LoadHtml(node.InnerHtml);

                    Project pro = new Project();
                    HtmlNode proNode = hd.DocumentNode.SelectSingleNode("/td[2]/a");
                    pro.ProjectID = StrHelper.GetRegexValue(proNode.Attributes["href"].Value, "&pjID=(?<1>\\d{1,20})&")[1];
                    pro.ProjectPresell = proNode.InnerText;
                    pro.ProjectName = hd.DocumentNode.SelectSingleNode("/td[3]").InnerText;
                    pro.ProjectUrl = ProjectDomain + proNode.Attributes["href"].Value;
                    //Console.WriteLine("ProjectPresell : " + pro.ProjectPresell);
                    //Console.WriteLine("ProjectName : " + pro.ProjectName);
                    listProject.Add(pro);
                }
                //Console.Read();
            }
            //Console.Read();
            return listProject;
        }
        public List<Project> getProjectList(String presellNo)
        {
            List<Project> listProject = new List<Project>();
            string html = UrlHelper.GetUrlHtml(ProjectDomain + "/search/project/projectSearch.jsp?presellNo=" + presellNo);
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//div[@id=\"main\"]/div[2]/table[4]/tr/td[2]");
            foreach (HtmlNode node in nodes)
            {
                Project pro = new Project();
                HtmlNode proNode = node.SelectSingleNode("a");
                pro.ProjectID = StrHelper.GetRegexValue(proNode.Attributes["href"].Value, "&pjID=(?<1>\\d{1,20})&")[1];
                pro.ProjectPresell = presellNo;          
                pro.ProjectName = node.InnerText;
                pro.ProjectUrl = ProjectDomain + proNode.Attributes["href"].Value;
                listProject.Add(pro);
            }
            return listProject;
        }

        public Project supplementProject(Project pro)
        {
            Console.WriteLine("    读取分析数据中......");
            pro.Estate = this.getEstate(pro);
            pro.BlockList = this.getBlockList(pro);
            pro.FlatList = this.getFlatList(pro);
            return pro;
        }

        private DB_ESTATE getEstate(Project pro)
        {
            DB_ESTATE dbEastate = new DB_ESTATE();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(UrlHelper.GetUrlHtml(pro.ProjectUrl));
            string html = doc.DocumentNode.SelectSingleNode("/html/head/script[3]").InnerHtml;

            ///基本信息
            string url_JBXX = this.ProjectDomain + "/search/project/project.jsp?pjID=" + StrHelper.GetRegexValue(html, "project.jsp\\?pjID=(?<1>\\d{1,100})")[1];
            doc = new HtmlDocument();
            doc.LoadHtml(UrlHelper.GetUrlHtml(url_JBXX));
            HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//table[@class=\"tab_style03\"][1]/tr/td[@class=\"tab_style01_td\"]");
            dbEastate.WYID = "1";
            dbEastate.CS = "广州";
            dbEastate.XZQ = nodes[4].InnerText;
            dbEastate.WYMC = nodes[0].InnerText;
            dbEastate.WYDZ = nodes[2].InnerText;
            dbEastate.ZDMJ = nodes[5].InnerText;
            dbEastate.ZJZMJ = nodes[6].InnerText;
            return dbEastate;
        }

        private List<DB_BLOCK> getBlockList(Project pro)
        {

            DB_BLOCK dbBlock = new DB_BLOCK();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(UrlHelper.GetUrlHtml(pro.ProjectUrl));
            string html = doc.DocumentNode.SelectSingleNode("/html/head/script[3]").InnerHtml;

            ///基本信息
            string url_JBXX = this.ProjectDomain + "/search/project/project.jsp?pjID=" + StrHelper.GetRegexValue(html, "project.jsp\\?pjID=(?<1>\\d{1,100})")[1];
            doc = new HtmlDocument();
            doc.LoadHtml(UrlHelper.GetUrlHtml(url_JBXX));
            HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//table[@class=\"tab_style03\"][1]/tr/td[@class=\"tab_style01_td\"]");
            dbBlock.WYID = pro.Estate.WYID;
            dbBlock.FZS = nodes[3].InnerText;
            //dbBlock.DDDLMCDZ = nodes[2].InnerText;

            ///施工许可证
            string[] agree = StrHelper.GetRegexValue(html, "workAgree.jsp\\?agreeName=(?<1>[^&]{0,100})&agreeId=(?<2>[0-9,]{0,100})");
            string url_SGXKZ = this.ProjectDomain + "/search/project/workAgree.jsp?agreeName=" + agree[1] + "&agreeId=" + agree[2];
            doc = new HtmlDocument();
            doc.LoadHtml(UrlHelper.GetUrlHtml(url_SGXKZ, "utf-8"));
            dbBlock.JGRQ = doc.DocumentNode.SelectSingleNode("//table[@class=\"tab_style03\"][1]/tr[8]/td[4]").InnerText;

            ///销控表
            string[] sellForm = StrHelper.GetRegexValue(html, "sellForm.jsp\\?pjID=(?<1>\\d{1,100})&presell=(?<2>\\d{1,100})&chnlname=(?<3>[a-z]{1,100})");
            string url_XKB = this.ProjectDomain + "/search/project/sellForm.jsp?pjID=" + sellForm[1] + "&presell=" + sellForm[2] + "&chnlname=" + sellForm[3];
            doc.LoadHtml(UrlHelper.GetUrlHtml(url_XKB));

            nodes = doc.DocumentNode.SelectNodes("//table[1]/tr[3]/td[1]/table[1]/tr/td");
            
            int iDLID = 0;
            List<DB_BLOCK> listBlock = new List<DB_BLOCK>();
            foreach (HtmlNode node in nodes)
            {
                if (node.InnerHtml.IndexOf("value=") > 0)
                {
                    iDLID++;
                    DB_BLOCK objBlock = new DB_BLOCK();
                    objBlock._BuildingID = StrHelper.GetRegexValue(node.InnerHtml, "value=\"(?<1>\\d{1,100})\"")[1];
                    objBlock.WYID = dbBlock.WYID;
                    objBlock.DLID = iDLID.ToString();
                    objBlock.FZS = dbBlock.FZS;
                    objBlock.JGRQ = dbBlock.JGRQ;
                    objBlock.LDZWMC = node.InnerText.Replace("&nbsp;", "");
                    listBlock.Add(objBlock);
                }
            }
            return listBlock;
        }


        private List<DB_FLAT> getFlatList(Project pro)
        {

            List<DB_FLAT> listFlat = new List<DB_FLAT>();

            foreach (DB_BLOCK objBP in pro.BlockList)
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(UrlHelper.GetUrlHtml(this.ProjectDomain + "/search/project/sellForm_pic.jsp?buildingID=" + objBP._BuildingID + "&chnlname=fdcxmxx"));

                HtmlNodeCollection nodes = doc.DocumentNode.SelectNodes("//table[1]/tr[position() > 2]//a");
                foreach (HtmlNode node in nodes)
                {
                    string title = node.Attributes["title"].Value.Replace("\r", "").Replace("\n", ""); ;
                    string[] arrTitle = StrHelper.GetRegexValue(title, "类型：(?<1>[^户]{0,10})户型：(?<2>[^套]{0,10})套内面积：(?<3>[^平]{1,10})平方米总面积：(?<4>[^平]{1,10})平方米");

                    DB_FLAT objFlat = new DB_FLAT();
                    objFlat.DLID = objBP.DLID;
                    objFlat.YTL1 = arrTitle[1];
                    string[] arrFT = StrHelper.GetRegexValue(arrTitle[2], "(?<1>[^房]{1})房(?<2>[^厅]{1})厅");
                    if (arrFT.Length > 0)
                    {
                        objFlat.F = getCHToInt(arrFT[1]);
                        objFlat.T = getCHToInt(arrFT[2]);
                    }
                    else
                    {
                        objFlat.F = "";
                        objFlat.T = "";
                    }
                    objFlat.TNMJ = arrTitle[3];
                    objFlat.JZMJ = arrTitle[4];
                    string dy = getDY(node.InnerText);
                    if (dy.Substring(0, 1) == "-")
                    {                        
                        objFlat.CSL = dy.Substring(0, 2);
                        objFlat.CSM = dy.Substring(0, 2);
                        objFlat.DY = dy.Substring(2);
                    }
                    else
                    {
                        objFlat.CSL = dy.Substring(0, dy.Length - 2);
                        objFlat.CSM = dy.Substring(0, dy.Length - 2);
                        objFlat.DY = dy.Substring(dy.Length - 2);
                    }

                    //doc = new HtmlDocument();
                    //doc.LoadHtml(UrlHelper.GetUrlHtml(this.ProjectDomain + "/search/project/"+ node.Attributes["href"].Value));
                    //HtmlNodeCollection detail = doc.DocumentNode.SelectNodes("//table[2]/tr[position() > 2]/td[@class=\"tab_style01_td\"]");
                    //if (detail.Count > 6)
                    //{
                    //    objFlat.CSL = detail[5].InnerText;
                    //    objFlat.CSM = detail[6].InnerText.Trim() == "" ? detail[5].InnerText : detail[6].InnerText;
                    //    string sFH = detail[3].InnerText.Trim() == "" ? detail[2].InnerText.Trim() : detail[3].InnerText.Trim();
                    //    objFlat.DY = sFH.Substring(objFlat.CSM.Length);
                    //    objFlat.YTL1 = detail[4].InnerText; 
                    //}
                    listFlat.Add(objFlat);
                }
            }
            return listFlat;
        }

        private string getNumber(string html)
        {
            string format = "<img src=\"/images/{0}.gif\" width=\"11\" height=\"12\">";
            html = StrHelper.ReplaceString(html, String.Format(format, "7af3fe7c42"), "0", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "73b40f64a0"), "1", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "4530c88c56"), "2", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "efd113682e"), "3", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "2300524172"), "4", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "e9df0cb5ec"), "5", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "8ad7bc8ae7"), "6", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "8e7390dc35"), "7", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "0ab21ea5fb"), "8", false);
            html = StrHelper.ReplaceString(html, String.Format(format, "77f1886353"), "9", false);
            return StrHelper.ReplaceString(html, "<br>", "", false);
        }
        
        private string getCHToInt(string strFT)
        {
            switch (strFT)
            {
                case "零":
                    strFT = "0";
                    break;
                case "一":
                    strFT = "1";
                    break;
                case "两":
                    strFT = "2";
                    break;
                case "二":
                    strFT = "2";
                    break;
                case "三":
                    strFT = "3";
                    break;
                case "四":
                    strFT = "4";
                    break;
                case "五":
                    strFT = "5";
                    break;
                case "六":
                    strFT = "6";
                    break;
                case "七":
                    strFT = "7";
                    break;
                case "八":
                    strFT = "8";
                    break;
                case "九":
                    strFT = "9";
                    break;
            }
            return strFT;
        }

        private string getDY(string dy)
        {
            return dy.Replace("★", "").Replace("☆", "").Replace("◎", "").Replace("■", "").Replace("◆", "").Trim();
        }


        public void SaveExcel(Project _Project)
        {
            Console.WriteLine("    保存数据至Excel中......");
            string strPath = @"Xls\";
            if (!Directory.Exists(strPath))
            {
                Directory.CreateDirectory(strPath);
            }
            strPath = strPath + _Project.ProjectPresell + "_"  + _Project.ProjectID + "_" + _Project.ProjectName + ".xls"; //保存的路径和文件名

            Stopwatch watch = new Stopwatch();
            watch.Start();
            ExcelWorkbook wbook = new ExcelWorkbook();
            wbook.SetDefaultFont("Arial", 10);
            ExcelWorksheet wsheet;
            ExcelCellStyle style;

            wbook.CreateSheet("DB_ESTATE");
            wsheet = wbook.GetSheet("DB_ESTATE");
            style = wbook.CreateStyle();
            style.Pattern = EnumFill.Solid;
            style.PatternForeColour = EnumColours.Grey25;
            style.Font.Size = 11;
            style.Font.Bold = true;
            string[] arrESTATE_Title = { "物业ID", "城市", "行政区", "行政区英文名称", "片区", "片区英文名称", "物业名称", "物业名称拼音索引", "物业名称英文名称", 
                                         "物业别名", "物业地址", "物业类型及规模", "环线位置", "周围环境", "楼盘概况", "占地面积", "总建筑面积", "土地使用分区", 
                                         "容积率", "绿化率", "项目规划", "建筑类型", "建筑设计", "园林设计", "车位数量", "目标客户群", "物业管理方式", 
                                         "核心卖点", "项目自身设施", "项目优劣势", "周边配套"};
            for (int o = 0; o < arrESTATE_Title.Length; o++)
            {
                wsheet.AddCell((ushort)(o + 1), 1, arrESTATE_Title[o], style);
            }
            DB_ESTATE dbEstate = _Project.Estate;
            wsheet.AddCell(1, 2, dbEstate.WYID);
            wsheet.AddCell(2, 2, dbEstate.CS);
            wsheet.AddCell(3, 2, dbEstate.XZQ);
            wsheet.AddCell(5, 2, dbEstate.XZQ);
            wsheet.AddCell(7, 2, dbEstate.WYMC);
            wsheet.AddCell(11, 2, dbEstate.WYDZ);
            wsheet.AddCell(16, 2, dbEstate.ZDMJ);
            wsheet.AddCell(17, 2, dbEstate.ZJZMJ);


            wbook.CreateSheet("DB_BLOCK");
            wsheet = wbook.GetSheet("DB_BLOCK");
            style = wbook.CreateStyle();
            style.Pattern = EnumFill.Solid;
            style.PatternForeColour = EnumColours.Grey25;
            style.Font.Size = 11;
            style.Font.Bold = true;
            string[] arrBLOCK_Title = { "物业ID", "大楼ID", "期数", "期数英文名称", "楼栋中文名称", "楼栋英文名称", 
                                         "楼栋别名1", "楼栋别名2", "座落位置", "楼栋结构", "竣工日期", "总单元数目", "建筑总面积", "用途(L1)", "用途(L2)",
                                         "车位类型", "户外车位", "有盖车位", "电梯数量", "扶梯数量", "售楼书号", "物业管理费", "物业管理费币值", "管理公司", "发展商", "单栋大楼名称地址", 
                                         "楼盘简述", "路/街/里/弄", "路/街/里/弄（英文）", "街号由", "街号由尾码", "街号至", "街号至尾码", "县", "镇","村", "房屋所有权证号", 
                                         "土地使用证证号", "地块编号", "土地使用权起始日", "土地使用权终止日", "土地使用权条款", "地上层数", "地下层数", "区域类别", "周围环境", "发展趋势", "楼栋概况", "楼梯及设施",
                                         "宗地号(报告相关)", "使用条款(报告相关)", "竣工日期(报告相关)", "市场可售性备注"};
            for (int o = 0; o < arrBLOCK_Title.Length; o++)
            {
                wsheet.AddCell((ushort)(o + 1), 1, arrBLOCK_Title[o], style);
            }
            ushort iRow = 1;
            foreach(DB_BLOCK block in _Project.BlockList.ToArray()){
                iRow++;
                wsheet.AddCell(1, iRow, block.WYID);
                wsheet.AddCell(2, iRow, block.DLID);
                wsheet.AddCell(5, iRow, block.LDZWMC);
                wsheet.AddCell(11, iRow, block.JGRQ);
                wsheet.AddCell(25, iRow, block.FZS);
            }

            wbook.CreateSheet("DB_FLAT");
            wsheet = wbook.GetSheet("DB_FLAT");
            style = wbook.CreateStyle();
            style.Pattern = EnumFill.Solid;
            style.PatternForeColour = EnumColours.Grey25;
            style.Font.Size = 10;
            style.Font.Bold = true;
            string[] strDBFLAT_Title = { "大楼ID", "层数列(数字)", "层数名", "单元", "单元名称", "建筑面积(平方米)", "套内面积(平方米)", "天井(平方米)",
                                         "露台(平方米)", "平台(平方米)", "天台(平方米)", "花园(平方米)", "阳台(平方米)", "窗台(平方米)", 
                                         "阁楼(平方米)", "储物室(平方米)", "空调室(平方米)", "用途(L1)", "用途(L2)", "房", "厅", "座向",
                                         "单元结构", "备注", "法定用途", "合并情况"};
            for (int o = 0; o < strDBFLAT_Title.Length; o++)
            {
                wsheet.AddCell((ushort)(o + 1), 1, strDBFLAT_Title[o], style);
            }
            iRow = 1;
            foreach (DB_FLAT flat in _Project.FlatList.ToArray())
            {
                iRow++;
                wsheet.AddCell(1, iRow, flat.DLID);
                wsheet.AddCell(2, iRow, flat.CSL);
                wsheet.AddCell(3, iRow, flat.CSM);
                wsheet.AddCell(4, iRow, flat.CSL + flat.DY);
                wsheet.AddCell(5, iRow, flat.DY);
                wsheet.AddCell(6, iRow, flat.JZMJ);
                wsheet.AddCell(7, iRow, flat.TNMJ);
                wsheet.AddCell(18, iRow, flat.YTL1);
                wsheet.AddCell(20, iRow, flat.F);
                wsheet.AddCell(21, iRow, flat.T);
            }

            wbook.Save(strPath);
            watch.Stop();


        }

    }
}
