using System;
using System.Web;
using System.Collections.Specialized;
using System.Text;
using System.Net;
using System.IO;

namespace DTZ.Helper
{

    /// <summary>
    /// URL常用方法  
    /// </summary>
    public class UrlHelper
    {
        private static string _encoding = "UTF-8";
        public static void SetEncoding(string sEncoding)
        {
            UrlHelper._encoding = sEncoding;
        }

        private static string _referer = "";
        public static void SetReferer(string sReferer)
        {
            UrlHelper._referer = sReferer;
        }


        /// <summary>
        /// 读取指定Url的Html源代码
        /// </summary>
        /// <param name="sUrl">指定Url</param>
        /// <param name="sEncoding">指定读取编码</param>
        /// <returns></returns>
        public static string GetUrlHtml(string sUrl)
        {
            return GetUrlHtml(sUrl, _encoding);
        }

        /// <summary>
        /// 读取指定Url的Html源代码
        /// </summary>
        /// <param name="sUrl">指定Url</param>
        /// <param name="sEncoding">指定读取编码</param>
        /// <returns></returns>
        public static string GetUrlHtml(string sUrl, string encoding)
        {
            Console.WriteLine("      " + sUrl);
            HttpWebRequest Req = (HttpWebRequest)WebRequest.Create(sUrl);
            Req.Referer = UrlHelper._referer;
            Req.Method = "GET";
            Req.UserAgent = "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1;)";
            Req.AllowAutoRedirect = true;
            Req.MaximumAutomaticRedirections = 10;
            // 超时时间30000=30秒
            Req.Timeout = 10000;
            //  是否建立TCP持久连接
            Req.KeepAlive = false;

            HttpWebResponse response = (HttpWebResponse)Req.GetResponse();
            Stream stream = response.GetResponseStream();
            Encoding myEncoding = Encoding.GetEncoding(encoding);
            StreamReader streamReader = new StreamReader(stream, myEncoding);
            string html = streamReader.ReadToEnd();
            streamReader.Close();
            streamReader.Dispose();
            stream.Close();
            stream.Dispose();
            response.Close();

            return html;
        }
    }
}
