using System;
using System.Xml;
using System.Web;

namespace DTZ.Helper
{
    /// <summary>
    /// 读取指定.config文件
    /// </summary>
    public class ConfigHelper
    {
        /// <summary>
        /// 非web程序默认文件名Parameters.config
        /// </summary>
        public static string DefaultFileName
        {
            get { return "Parameters.config"; }
        }
        /// <summary>
        /// 读取指定.config文件中的Key值
        /// </summary>
        /// <param name="Key"></param>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public static string AppSettings(string Key, string FileName)
        {
            return ReadValue(Key, FileName);
        }
        /// <summary>
        /// 读取指定Key的值
        ///  （web程序中读取web.config 非web程序中读取Parameters.config ）
        /// </summary>
        /// <param name="Key"></param>
        /// <returns></returns>
        public static string AppSettings(string Key)
        {
            if (HttpContext.Current != null)
            {
                return System.Configuration.ConfigurationManager.AppSettings[Key]; ;
            }
            else //非web程序引用
            {
                return ReadValue(Key, DefaultFileName);
            }            
        }

        private static string ReadValue(string Key, string FileName)
        {
            string appSettingValue = "";
            using (XmlTextReader xmlTR = new XmlTextReader(FileName))
            {
                while (xmlTR.Read())
                {
                    if (xmlTR.NodeType == XmlNodeType.Element)
                    {
                        if (xmlTR.Name.ToLower() == "add")
                        {
                            if (xmlTR.GetAttribute("key").ToLower() == Key.ToLower())
                            {
                                appSettingValue = xmlTR.GetAttribute("value");
                                break;
                            }
                        }
                    }
                }
                xmlTR.Close();
            }
            return appSettingValue;
        }
    }
}
