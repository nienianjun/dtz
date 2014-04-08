using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml.Serialization;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;

namespace DTZ.Helper
{
    /// <summary>
    /// 实体类转换
    /// </summary>
    public class DataModelHelper
    {
        ///// <summary>
        ///// 实体类转换成DataTable
        ///// </summary>
        ///// <param name="arrDataModel">实体类</param>
        ///// <returns></returns>
        //public static DataTable ToDataTable(Object[] arrDataModel)
        //{
        //    DataTable dt = new DataTable();
        //    if (arrDataModel.Length > 0)
        //    {
        //        Type type = arrDataModel[0].GetType();

        //        System.Reflection.FieldInfo[] arrInfo = type.GetFields();

        //        foreach (System.Reflection.FieldInfo info in arrInfo)
        //        {
        //            dt.Columns.Add(new DataColumn(info.Name, info.FieldType));
        //        }
        //        for (int i = 0; i <= arrDataModel.Length - 1; i++)
        //        {
        //            if (arrDataModel[i] != null)
        //            {
        //                dt.Rows.Add(dt.NewRow());
        //                for (int j = 0; j <= dt.Columns.Count - 1; j++)
        //                {
        //                    dt.Rows[i].ItemArray[j] = arrInfo[j].GetValue(arrDataModel[i]);
        //                }
        //            }
        //        }
        //    }
        //    return dt;
        //}

        /// <summary>
        /// 实体类转换成XML字符串[旧方法，请使用ToXML]
        ///   string | String[] | Int32 | Int32[] | double | double[] | DateTime
        /// </summary>
        /// <param name="arrDataModel">实体类</param>
        /// <returns></returns>
        public static String ToXmlString(Object[] arrDataModel)
        {
            StringBuilder sbXML = new StringBuilder();
            if (arrDataModel.Length > 0)
            {
                Type type = arrDataModel[0].GetType();
                System.Reflection.FieldInfo[] arrInfo = type.GetFields();

                for (int i = 0; i <= arrDataModel.Length - 1; i++)
                {
                    sbXML.AppendLine(string.Format("<{0}>", type.Name));
                    if (arrDataModel[i] != null)
                    {
                        foreach (System.Reflection.FieldInfo info in arrInfo)
                        {
                            if (info.GetValue(arrDataModel[i]) != null)
                            {
                                Type infoType = info.FieldType;

                                if (infoType.Equals(typeof(string)))
                                {
                                    if (info.GetValue(arrDataModel[i]).ToString().Trim() == "")
                                        sbXML.AppendLine(string.Format("<{0}/>", info.Name));
                                    else
                                        sbXML.AppendLine(string.Format("<{0}><![CDATA[ {1} ]]></{0}>", info.Name, info.GetValue(arrDataModel[i]).ToString()));
                                }
                                else if (infoType.Equals(typeof(String[])))
                                {
                                    sbXML.AppendLine(string.Format("<{0}>", info.Name));
                                    foreach (string sValue in (string[])info.GetValue(arrDataModel[i]))
                                    {
                                        if (sValue.Trim() == "")
                                            sbXML.AppendLine(string.Format("<string></string>"));
                                        else
                                            sbXML.AppendLine(string.Format("<string><![CDATA[ {0} ]]></string>", sValue));
                                    }
                                    sbXML.AppendLine(string.Format("</{0}>", info.Name));
                                }
                                else if (infoType.Equals(typeof(Nullable<Int32>)) || infoType.Equals(typeof(Int32)) || infoType.Equals(typeof(Nullable<double>)) || infoType.Equals(typeof(double)))
                                {
                                    sbXML.AppendLine(string.Format("<{0}>{1}</{0}>", info.Name, info.GetValue(arrDataModel[i]).ToString()));
                                }
                                else if (infoType.Equals(typeof(DateTime)))
                                {
                                    sbXML.AppendLine(string.Format("<{0}>{1}</{0}>", info.Name, ((DateTime)info.GetValue(arrDataModel[i])).ToString()));
                                }
                                else if (infoType.Equals(typeof(Int32[])))
                                {
                                    sbXML.AppendLine(string.Format("<{0}>", info.Name));
                                    foreach (Int32 iValue in (Int32[])info.GetValue(arrDataModel[i]))
                                    {
                                        sbXML.AppendLine(string.Format("<int>{0}</int>", iValue));
                                    }
                                    sbXML.AppendLine(string.Format("</{0}>", info.Name));
                                }
                                else if (infoType.Equals(typeof(double[])))
                                {
                                    sbXML.AppendLine(string.Format("<{0}>", info.Name));
                                    foreach (double iValue in (double[])info.GetValue(arrDataModel[i]))
                                    {
                                        sbXML.AppendLine(string.Format("<double><![CDATA[ {0} ]]></double>", iValue));
                                    }
                                    sbXML.AppendLine(string.Format("</{0}>", info.Name));
                                }
                                else
                                {
                                    sbXML.AppendLine(string.Format("<{0}>", info.Name));
                                    try
                                    {
                                        sbXML.AppendLine(ToXmlString((object[])info.GetValue(arrDataModel[i])));
                                    }
                                    catch
                                    {
                                        sbXML.AppendLine(ToXmlString(new object[] { (object)info.GetValue(arrDataModel[i]) }));
                                    }
                                    sbXML.AppendLine(string.Format("</{0}>", info.Name));
                                }
                            }
                            //else
                            //{
                            //    if (blNullNode)
                            //        sbXML.AppendLine(string.Format("<{0}/>", info.Name));
                            //}
                        }
                    }
                    sbXML.AppendLine(string.Format("</{0}>", type.Name));
                }
            }
            return sbXML.ToString().Trim();

        }

        /// <summary>
        /// 实体类转换成XML字符串
        /// </summary>
        /// <typeparam name="T">实体类型名</typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ToXML<T>(T obj)
        {
            if (obj == null) return null;
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            MemoryStream stream = new MemoryStream();
            XmlTextWriter xtw = new XmlTextWriter(stream, Encoding.UTF8);
            xtw.Formatting = Formatting.Indented;
            try
            {
                serializer.Serialize(stream, obj);
            }
            catch { return ""; }

            stream.Position = 0;
            StringBuilder sbXML = new StringBuilder();
            using (StreamReader sr = new StreamReader(stream, Encoding.UTF8))
            {
                sr.ReadLine();
                sbXML.AppendLine(sr.ReadToEnd().Replace(" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"", "").Replace(" xsi:nil=\"true\"", ""));
                sr.Close();
                sr.Dispose();
            }
            stream.Close();
            stream.Dispose();
            return sbXML.ToString();
        }

        /// <summary>
        /// 实体类数组转换成XML字符串
        /// </summary>
        /// <typeparam name="T">实体类型名</typeparam>
        /// <param name="obj"></param>
        /// <param name="sNodeName">数组外层标签替换名</param>
        /// <returns></returns>
        public static string ToXML<T>(T obj, string sNodeName)
        {
            string sTypeName = obj.GetType().Name.Replace("[]", "");
            
            string sXML = ToXML<T>(obj);
            sXML = Regex.Replace(sXML, "<ArrayOf" + sTypeName + ">", "<" + sNodeName + ">", RegexOptions.IgnoreCase);
            sXML = Regex.Replace(sXML, "</ArrayOf" + sTypeName + ">", "</" + sNodeName + ">", RegexOptions.IgnoreCase);
            sXML = Regex.Replace(sXML, "<ArrayOf" + sTypeName + "/>", "<" + sNodeName + "/>", RegexOptions.IgnoreCase);
            return sXML;
        }   

        /// <summary>
        /// 实体类转换成JSON字符串 (System.Web.Script.Serialization.JavaScriptSerializer())
        /// </summary>
        /// <param name="objDataModel">实体类数组</param>
        /// <returns></returns>
        public static String ToJSON(Object objDataModel)
        {
            System.Web.Script.Serialization.JavaScriptSerializer JSS = new System.Web.Script.Serialization.JavaScriptSerializer();
            return JSS.Serialize(objDataModel);
        }
        /// <summary>
        /// 实体类数组转换成JSON字符串 (System.Web.Script.Serialization.JavaScriptSerializer())
        /// </summary>
        /// <param name="arrDataModel">实体类数组</param>
        /// <returns></returns>
        public static String ToJSON(Object[] arrDataModel)
        {
            System.Web.Script.Serialization.JavaScriptSerializer JSS = new System.Web.Script.Serialization.JavaScriptSerializer();
            return JSS.Serialize(arrDataModel);
        }




        /// <summary>
        /// XML到反序列化到对象----支持泛类型
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sXML"></param>
        /// <returns></returns>
        public static T ToObjFromXML<T>(string sXML)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                using (StreamWriter sw = new StreamWriter(stream, Encoding.UTF8))
                {
                    sw.Write(sXML);
                    sw.Flush();
                    stream.Seek(0, SeekOrigin.Begin);
                    XmlSerializer serializer = new XmlSerializer(typeof(T));
                        return ((T)serializer.Deserialize(stream));
                }
            }
        }

        /// <summary>
        /// JSON到反序列化到对象  (System.Web.Script.Serialization.JavaScriptSerializer())
        /// </summary>
        /// <param name="sJSON">JSON</param>
        /// <returns></returns>
        public static T ToObjFromJSON<T>(string sJSON)
        {
            System.Web.Script.Serialization.JavaScriptSerializer JSS = new System.Web.Script.Serialization.JavaScriptSerializer();
            return JSS.Deserialize<T>(sJSON);
        }

        //public static String ToXmlString(Object[] arrDataModel)
        //{
        //    string sXMLFormat = "<{0}>{1}</{0}>";
        //    string sXMLValue = "";

        //    StringBuilder sbXML = new StringBuilder();
        //    if (arrDataModel.Length > 0)
        //    {
        //        Type type = arrDataModel[0].GetType();
        //        System.Reflection.FieldInfo[] arrInfo = type.GetFields();

        //        for (int i = 0; i <= arrDataModel.Length - 1; i++)
        //        {
        //            sbXML.AppendLine(string.Format("<{0}>", type.Name));
        //            if (arrDataModel[i] != null)
        //            {
        //                foreach (System.Reflection.FieldInfo info in arrInfo)
        //                {
        //                    if (info.GetValue(arrDataModel[i]) != null)
        //                    {
        //                        sXMLValue = info.GetValue(arrDataModel[i]).ToString();
        //                        if (sXMLValue.Length > 0)
        //                        {
        //                            if (info.FieldType.ToString() == "System.String")
        //                                sXMLValue = String.Format("<![CDATA[ {0} ]]>", sXMLValue);
        //                            sbXML.AppendLine(string.Format(sXMLFormat, info.Name, sXMLValue));
        //                        }
        //                        else
        //                            sbXML.AppendLine("<" + info.Name + "/>");
        //                    }
        //                    else
        //                        sbXML.AppendLine("<" + info.Name + "/>");
        //                }
        //            }
        //            sbXML.AppendLine(string.Format("</{0}>", type.Name));
        //        }
        //    }
        //    return sbXML.ToString();
        //}



    }
}
