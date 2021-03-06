using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Lau.Net.Utils
{
    public class SerializationUtil
    {

        #region XML序列化和反序列化
        /// <summary>
        /// 对象序列化成 XML String
        /// </summary>
        public static byte[] XmlSerialize<T>(T data)
        {
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
                using (MemoryStream ms = new MemoryStream())
                {
                    xmlSerializer.Serialize(ms, data);
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        /// <summary>
        /// XML String 反序列化成对象
        /// </summary>
        public static T XmlDeserialize<T>(byte[] data)
        {
            T t = default(T);
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
                using (Stream xmlStream = new MemoryStream(data))
                {
                    using (XmlReader xmlReader = XmlReader.Create(xmlStream))
                    {
                        Object obj = xmlSerializer.Deserialize(xmlReader);
                        t = (T)obj;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return t;
        }

        #endregion


        #region 二进制序列化和反序列化
        /// <summary>
        /// 二进制序列化
        /// </summary>
        public static byte[] BinarySerialize<T>(T data)
        {
            try
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                using (MemoryStream ms = new MemoryStream())
                {
                    binaryFormatter.Serialize(ms, data);
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 二进制反序列化
        /// </summary>
        public static T BinaryDeserialize<T>(byte[] data)
        {
            BinaryFormatter binaryFormatter = new BinaryFormatter();
            using (MemoryStream ms = new MemoryStream(data))
            {
                return (T)binaryFormatter.Deserialize(ms);
            }
        }

        #endregion

        #region Soap序列化和反序列化
        //public static byte[] SoapSerialize<T>(T data)
        //{
        //    try
        //    {
        //        SoapFormatter soapFormatter = new SoapFormatter();
        //        using (MemoryStream ms = new MemoryStream())
        //        {
        //            soapFormatter.Serialize(ms, data);
        //            return ms.ToArray();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(ex.Message);
        //    }
        //}

        //public static T SoapDeserialize<T>(byte[] data)
        //{
        //    SoapFormatter soapFormatter = new SoapFormatter();
        //    using (MemoryStream ms = new MemoryStream(data))
        //    {
        //        return (T)soapFormatter.Deserialize(ms);
        //    }
        //}
        #endregion

    }
}
