using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Lau.Net.Utils
{
    /// <summary>
    /// XML 文档中的每个成分都是一个节点。
    ///•整个文档是一个文档节点
    ///•每个 XML 标签是一个元素节点
    ///•包含在 XML 元素中的文本是文本节点
    ///•每一个 XML 属性是一个属性节点
    ///•注释属于注释节点 
    /// </summary>
    public class XmlUtil
    {
        private string m_XmlFilePath;
        private XmlDocument m_XmlDoc = new XmlDocument();

        /// <summary>
        /// 构造函数（加载XML文件)
        /// </summary>
        /// <param name="xmlFilePath">XML文件路径</param> 
        public XmlUtil(string xmlFilePath)
        {
            try
            {
                this.m_XmlFilePath = xmlFilePath;
                this.m_XmlDoc.Load(xmlFilePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 构造函数（新建XML文件）
        /// </summary>
        /// <param name="xmlFileSavePath">文件保存路径(含文件名)</param>
        /// <param name="rootName">根节点名称</param> 
        public XmlUtil(string xmlFileSavePath, string rootName)
        {
            try
            {
                this.m_XmlFilePath = xmlFileSavePath;
                XmlDeclaration decl = this.m_XmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
                this.m_XmlDoc.AppendChild(decl);
                XmlNode newNode = m_XmlDoc.CreateElement(rootName);
                this.m_XmlDoc.AppendChild(newNode);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region 查找
        /// <summary>
        /// 取得指定XPath表达式的节点对象
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param>
        /// <returns>返回节点对象</returns>
        public XmlNode GetNode(string xmlPath)
        {
            return m_XmlDoc.SelectSingleNode(xmlPath);
        }

        /// <summary>
        /// 取得指定XPath表达式的节点对象下的子节点列表
        /// </summary>
        /// <param name="xPath"></param>
        /// <returns></returns>
        public XmlNodeList GetNodeList(string xPath)
        {
            XmlNodeList nodeList = m_XmlDoc.SelectSingleNode(xPath).ChildNodes;
            return nodeList;

        }

        /// <summary>
        /// 取得指定XPath表达式的节点对象的InnerText
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param>
        /// <returns>返回节点的值</returns>
        public string GetNodeInnerText(string xmlPath)
        {
            return this.GetNode(xmlPath).InnerText;
        }

        /// <summary>
        /// 取得指定XPath表达式的节点对象的属性值
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param>
        /// <param name="attrName">属性名</param>
        /// <returns>返回属性值</returns>
        public string GetNodeAttrbuteValue(string xmlPath, string attrName)
        {
            return this.GetNode(xmlPath).Attributes[attrName].Value;
        }

        /// <summary>
        /// 取得指定XPath表达式的元素节点对象
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param>
        /// <returns>返回元素节点对象</returns>
        public XmlElement GetElement(string xmlPath)
        {
            try
            {
                return (XmlElement)this.GetNode(xmlPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region 修改
        /// <summary>
        /// 修改指定XPath表达式的元素节点对象的InnerText
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param>
        /// <param name="innerText">修改后的文本</param>
        public void UpdateElementInnerText(string xmlPath, string innerText)
        {
            this.GetElement(xmlPath).InnerText = innerText;
        }

        /// <summary>
        /// 修改指定XPath表达式的元素节点对象的属性值
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param> 
        /// <param name="attrName">属性名</param>
        /// <param name="attrValue">属性值</param>
        public void UpdateElementAttributeValue(string xmlPath, string attrName, string attrValue)
        {
            this.GetElement(xmlPath).Attributes[attrName].Value = attrValue;
            //this.GetElement(xmlPath).SetAttribute(attrName, attrValue);
        }
        #endregion

        #region 删除
        /// <summary>
        /// 删除指定XPath表达式的元素节点对象的属性
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param> 
        /// <param name="attrName">属性名</param> 
        public void DeleteElementAttribute(string xmlPath, string attrName)
        {
            this.GetElement(xmlPath).RemoveAttribute(attrName);
        }

        /// <summary>
        /// 删除指定XPath表达式的元素节点对象
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param> 
        public void DeleteElement(string xmlPath)
        {
            XmlNode node = this.GetNode(xmlPath);
            node.ParentNode.RemoveChild(node);
        }
        #endregion

        #region 添加
        /// <summary>
        /// 插入元素节点，并设置该元素的属性名称及属性值
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param>
        /// <param name="elementName">元素节点名称</param>
        /// <param name="attrName">属性名称</param>
        /// <param name="attrValue">属性值</param>
        /// <returns>返回插入的元素</returns>
        public XmlElement InsertElementNode(string xmlPath, string elementName, string attrName, string attrValue)
        {
            XmlElement element = this.InsertElementNode(xmlPath, elementName);
            element.SetAttribute(attrName, attrValue);
            return element;
        }

        /// <summary>
        /// 插入元素节点，并设置该元素的文本节点
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param>
        /// <param name="elementName">元素节点名称</param>
        /// <param name="innerText">文本内容</param>
        /// <returns>返回插入的元素</returns>
        public XmlElement InsertElementNode(string xmlPath, string elementName, string innerText)
        {
            XmlElement element = this.InsertElementNode(xmlPath, elementName);
            element.InnerText = innerText;
            return element;
        }

        /// <summary>
        /// 插入元素节点，并返回该元素
        /// </summary>
        /// <param name="xmlPath">要匹配的XPath表达式(例如:"根节点/节点名/子节点名)</param>
        /// <param name="elementName">元素节点名称</param>
        /// <returns>返回插入的元素</returns>
        public XmlElement InsertElementNode(string xmlPath, string elementName)
        {
            XmlNode parentNode = this.GetNode(xmlPath);    //查找指定路径的节点
            XmlElement element = m_XmlDoc.CreateElement(elementName);    //创建一个元素
            parentNode.AppendChild(element);
            return element;
        }

        /// <summary>
        /// 为指定元素添加属性
        /// </summary>
        /// <param name="element">需要添加属性的元素</param>
        /// <param name="attrName">属性名称</param>
        /// <param name="attrValue">属性值</param>
        /// <returns></returns>
        public XmlUtil AddAttribute(XmlElement element, string attrName, string attrValue)
        {
            element.SetAttribute(attrName, attrValue);
            return this;
        }

        /// <summary>
        /// 为指定元素设置文本节点
        /// </summary>
        /// <param name="element">要设置文本节点的元素</param>
        /// <param name="innerText">文本内容</param>
        /// <returns></returns>
        public XmlUtil AddInnerText(XmlElement element, string innerText)
        {
            element.InnerText = innerText;
            return this;
        }
        #endregion

        /// <summary>
        /// 保存
        /// </summary>
        public void Save()
        {
            try
            {
                this.m_XmlDoc.Save(m_XmlFilePath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        ///// <summary> 
        ///// 读取xml文件，并将文件序列化为类 
        ///// </summary> 
        ///// <typeparam name="T"></typeparam> 
        ///// <param name="path"></param> 
        ///// <returns></returns> 
        //public static T ReadXML<T>(string path)
        //{ 
        //    XmlSerializer reader = new XmlSerializer(typeof(T)); 
        //    StreamReader file = new StreamReader(@path); 
        //    return (T)reader.Deserialize(file); 
        //}

        //#region 读取XML资源到DataSet中 
        ///// <summary> 
        ///// 读取XML资源到DataSet中 
        ///// </summary> 
        ///// <param name="source">XML资源，文件为路径，否则为XML字符串</param> 
        ///// <param name="xmlType">XML资源类型</param> 
        ///// <returns>DataSet</returns> 
        //public static DataSet GetDataSet(string source, XmlType xmlType)
        //{ 
        //    DataSet ds = new DataSet(); 
        //    if (xmlType == XmlType.File)
        //    { 
        //        ds.ReadXml(source); 
        //    } 
        //    else
        //    { 
        //        XmlDocument xd = new XmlDocument(); 
        //        xd.LoadXml(source); 
        //        XmlNodeReader xnr = new XmlNodeReader(xd); 
        //        ds.ReadXml(xnr); 
        //    } 
        //    return ds; 
        //} 
        //#endregion



        ///// <summary>
        /////  读取节点中某一个属性的值。如果attribute为空，则返回整个节点的InnerText，否则返回具体attribute的值
        ///// </summary>
        ///// <param name="xmlPathNode">要匹配的XPath表达式(例如:"/节点名/子节点名</param>
        ///// <param name="attribute">节点中的属性</param>
        ///// <returns>如果attribute为空，则返回整个节点的InnerText，否则返回具体attribute的值</returns>
        ///// 使用实例: Read( "PersonF/person[@Name='Person2']", "");
        /////           Read( "PersonF/person[@Name='Person2']", "Name");
        //public string Read(string xmlPathNode, string attribute)
        //{
        //    string value = "";
        //    try
        //    {
        //        XmlNode xn = this.m_XmlDoc.SelectSingleNode(xmlPathNode);
        //        value = (attribute.Equals("") ? xn.InnerText : xn.Attributes[attribute].Value);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //    return value;
        //}

    }
}
