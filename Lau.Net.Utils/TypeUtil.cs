using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Lau.Net.Utils
{
    public static class TypeUtil
    {
        /// <summary>
        /// 判断列表中是否存在项
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static bool HasItem(this IEnumerable<object> list)
        {
            if (list != null && list.Any())
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 通用类型转换方法，EG:"".As<String>()
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T As<T>(this object obj)
        {
            Type type = typeof(T);
            if (typeof(T) == typeof(string))
            {
                return (T)Convert.ChangeType(obj + "", type);
            }
            try
            {
                if (type.Name.ToLower() == "guid")
                {
                    return (T)(object)new Guid(obj.ToString());
                }
                if (obj is IConvertible)
                {
                    return (T)Convert.ChangeType(obj, type);
                }
                if (obj == null)
                {
                    return default(T);
                }
                return (T)obj;
            }
            catch
            {
                return default(T);
            }
        }

        #region 代码被注释
        //#region 基础类型转换
        ///// <summary>
        /////  转换为布尔类型，为Yes时同样返回True
        ///// </summary>
        ///// <param name="obj"></param>
        ///// <returns>obj为空或转换失败时返回false</returns>
        //public static Boolean ToBoolean(object obj)
        //{
        //    if (obj == DBNull.Value || obj == null) return false;
        //    if (obj is bool) return (bool)obj;
        //    if (string.Equals(obj.ToString().ToLower(), "yes")) return true;
        //    if (string.Equals(obj.ToString().ToLower(), "no")) return false;
        //    try
        //    {
        //        return Convert.ToBoolean(obj);
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        ///// <summary>
        ///// 增加了对Obj为DBNull的判断
        ///// </summary>
        ///// <param name="obj"></param>
        ///// <returns></returns>
        //public static string ToString(object obj)
        //{
        //    if (obj == DBNull.Value || obj == null) return string.Empty;
        //    if (obj is string) return (string)obj;
        //    return obj.ToString();
        //}

        ///// <summary>
        ///// 如果obj为DBNull，则返回null
        ///// </summary>
        ///// <param name="obj"></param>
        ///// <returns></returns>
        //public static object ToObject(object obj)
        //{
        //    if (obj == DBNull.Value) return null;
        //    return obj;
        //}

        ///// <summary>
        ///// 转换为decimal类型
        ///// </summary>
        ///// <param name="obj"></param>
        ///// <param name="defaultValue">为空或转换失败时返回的值</param>
        ///// <returns>obj为空或转换失败时返回defaultValue</returns>
        //public static decimal ToDecimal(object obj, decimal defaultValue = 0)
        //{
        //    if (obj == DBNull.Value || obj == null) return defaultValue;
        //    try
        //    {
        //        return Convert.ToDecimal(obj);
        //    }
        //    catch
        //    {
        //        return defaultValue;
        //    }

        //}

        ///// <summary>
        ///// 转换为Int32(int)类型
        ///// </summary>
        ///// <param name="obj"></param>
        ///// <param name="defaultValue">为空或转换失败时返回的值</param>
        ///// <returns>obj为空或转换失败时返回defaultValue</returns>
        //public static int ToInt(object obj, int defaultValue = 0)
        //{
        //    if (obj == DBNull.Value || obj == null) return defaultValue;
        //    try
        //    {
        //        return Convert.ToInt32(obj);
        //    }
        //    catch
        //    {
        //        return defaultValue;
        //    }
        //}

        ///// <summary>
        ///// 转换为Int16类型
        ///// </summary>
        ///// <param name="obj"></param>
        ///// <param name="defaultValue">为空或转换失败时返回的值</param>
        ///// <returns>obj为空或转换失败时返回defaultValue</returns>
        //public static Int16 ToInt16(object obj, Int16 defaultValue = 0)
        //{
        //    if (obj == DBNull.Value || obj == null) return defaultValue;
        //    try
        //    {
        //        return Convert.ToInt16(obj);
        //    }
        //    catch
        //    {
        //        return defaultValue;
        //    }
        //}

        ///// <summary>
        ///// 转换为日期字符串
        ///// </summary>
        ///// <param name="obj"></param>
        ///// <param name="format">转换格式</param>
        ///// <returns>obj为空或转换失败时返回空字符串</returns>
        //public static string ToDateTimeString(object obj, string format = "yyyy/MM/dd")
        //{
        //    if (obj == DBNull.Value || obj == null) return string.Empty;
        //    try
        //    {
        //        return Convert.ToDateTime(obj).ToString(format);
        //    }
        //    catch
        //    {
        //        return string.Empty;
        //    }
        //}

        ///// <summary>
        ///// 转换为DateTime类型，obj为空或转换失败时返回DateTime.MinValue
        ///// </summary>
        ///// <param name="obj"></param>
        ///// <returns>obj为空或转换失败时返回DateTime.MinValue</returns>
        //public static DateTime ToDateTime(object obj)
        //{
        //    if (obj == DBNull.Value || obj == null) return DateTime.MinValue;
        //    try
        //    {
        //        return Convert.ToDateTime(obj);
        //    }
        //    catch
        //    {
        //        return DateTime.MinValue;
        //    }
        //}
        //#endregion

        //#region 两个对象比较大小
        ///// <summary>
        ///// 比较两个日期的大小（只比较日期部分)
        ///// 注：第一个时间大于第二个时间则返回1
        ///// </summary>
        ///// <param name="date1"></param>
        ///// <param name="date2"></param>
        ///// <returns>等于返回0，date1大于date2则返回1，反之则返回-1</returns>
        //public static int CompareOfDate(object date1, object date2)
        //{
        //    try
        //    {
        //        DateTime d1 = Convert.ToDateTime(ToDateTimeString(date1));
        //        DateTime d2 = Convert.ToDateTime(ToDateTimeString(date2));
        //        return DateTime.Compare(d1, d2);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        ///// <summary>
        ///// 比较两个字符串版本号的大小
        ///// </summary>
        ///// <param name="version1"></param>
        ///// <param name="version2"></param>
        ///// <returns>等于返回0，version1大于version2返回1，反之返回-1</returns>
        //public static int CompareOfVersion(string version1, string version2)
        //{
        //    try
        //    {
        //        Version v1 = new Version(version1);
        //        Version v2 = new Version(version2);
        //        return v1.CompareTo(v2);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}
        //#endregion

        //#region 常用进制转换
        ///// <summary>
        ///// 十六进制转二进制
        ///// </summary>
        ///// <param name="hex">十六进制字符串</param>
        ///// <returns>返回16进制对应的十进制数值</returns>
        //public static int Hex2Dec(string hex)
        //{
        //    return Convert.ToInt32(hex, 16);
        //}

        ///// <summary>
        ///// 十六进制转二进制
        ///// </summary>
        ///// <param name="hex">16进制字符串</param>
        ///// <returns>返回16进制对应的2进制字符串</returns>
        //public static string Hex2Bin(string hex)
        //{
        //    return Dec2Bin(Hex2Dec(hex));
        //}

        ///// <summary>
        ///// 十六进制转字节数组
        ///// </summary>
        ///// <param name="hex">16进制字符串</param>
        ///// <returns>返回16进制对应的字节数组</returns>
        //public static byte[] Hex2Bytes(string hex)
        //{
        //    MatchCollection mc = Regex.Matches(hex, @"(?i)[\da-f]{2}");
        //    return (from Match m in mc select Convert.ToByte(m.Value, 16)).ToArray();

        //    //hexString = hexString.Replace(" ", "");
        //    //if ((hexString.Length % 2) != 0)
        //    //    hexString += " ";
        //    //byte[] returnBytes = new byte[hexString.Length / 2];
        //    //for (int i = 0; i < returnBytes.Length; i++)
        //    //    returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
        //    //return returnBytes;   
        //}

        ///// <summary>
        ///// 十进制转二进制
        ///// </summary>
        ///// <param name="value"></param>
        ///// <returns></returns>
        //public static string Dec2Bin(int value)
        //{
        //    return Convert.ToString(value, 2);
        //}

        ///// <summary>
        ///// 十进制转十六进制
        ///// </summary>
        ///// <param name="value"></param>
        ///// <returns></returns>
        //public static string Dec2Hex(int value)
        //{
        //    return Convert.ToString(value, 16).ToUpper();
        //}

        ///// <summary>
        ///// 十进制转十六进制:格式化十六进制为指定位数，不足位数左边补0
        ///// </summary>
        ///// <param name="value">十进制数值</param>
        ///// <param name="formatLength">十六进制结果的总长度</param>
        ///// <returns>返回10进制数值对应的指定长度的16进制字符串</returns>
        //public static string Dec2Hex(int value, int formatLength)
        //{
        //    string hex = Dec2Hex(value);
        //    if (hex.Length >= formatLength) return hex;
        //    return hex.PadLeft(formatLength, '0');
        //}
        //#endregion

        //#region SqlDbType SqlTypeStringToSqlDbType(string sqlTypeString)
        ///// <summary>
        ///// sql server数据类型字符串（如：varchar）转换为SqlDbType类型
        ///// </summary>
        ///// <param name="sqlTypeString">sql server数据类型字符串</param>
        ///// <returns>SqlDbType</returns>
        //public static SqlDbType SqlTypeStringToSqlDbType(string sqlTypeString)
        //{
        //    SqlDbType dbType = SqlDbType.Variant;//默认为Object

        //    switch (sqlTypeString.ToLower())
        //    {
        //        case "int":
        //            dbType = SqlDbType.Int;
        //            break;
        //        case "varchar":
        //            dbType = SqlDbType.VarChar;
        //            break;
        //        case "bit":
        //            dbType = SqlDbType.Bit;
        //            break;
        //        case "datetime":
        //            dbType = SqlDbType.DateTime;
        //            break;
        //        case "decimal":
        //            dbType = SqlDbType.Decimal;
        //            break;
        //        case "float":
        //            dbType = SqlDbType.Float;
        //            break;
        //        case "image":
        //            dbType = SqlDbType.Image;
        //            break;
        //        case "money":
        //            dbType = SqlDbType.Money;
        //            break;
        //        case "ntext":
        //            dbType = SqlDbType.NText;
        //            break;
        //        case "nvarchar":
        //            dbType = SqlDbType.NVarChar;
        //            break;
        //        case "smalldatetime":
        //            dbType = SqlDbType.SmallDateTime;
        //            break;
        //        case "smallint":
        //            dbType = SqlDbType.SmallInt;
        //            break;
        //        case "text":
        //            dbType = SqlDbType.Text;
        //            break;
        //        case "bigint":
        //            dbType = SqlDbType.BigInt;
        //            break;
        //        case "binary":
        //            dbType = SqlDbType.Binary;
        //            break;
        //        case "char":
        //            dbType = SqlDbType.Char;
        //            break;
        //        case "nchar":
        //            dbType = SqlDbType.NChar;
        //            break;
        //        case "numeric":
        //            dbType = SqlDbType.Decimal;
        //            break;
        //        case "real":
        //            dbType = SqlDbType.Real;
        //            break;
        //        case "smallmoney":
        //            dbType = SqlDbType.SmallMoney;
        //            break;
        //        case "sql_variant":
        //            dbType = SqlDbType.Variant;
        //            break;
        //        case "timestamp":
        //            dbType = SqlDbType.Timestamp;
        //            break;
        //        case "tinyint":
        //            dbType = SqlDbType.TinyInt;
        //            break;
        //        case "uniqueidentifier":
        //            dbType = SqlDbType.UniqueIdentifier;
        //            break;
        //        case "varbinary":
        //            dbType = SqlDbType.VarBinary;
        //            break;
        //        case "xml":
        //            dbType = SqlDbType.Xml;
        //            break;
        //    }
        //    return dbType;
        //}
        //#endregion

        //#region Type SqlTypeToCsharpType(SqlDbType sqlDbType)
        ///// <summary>
        ///// SqlDbType转换为C#数据类型
        ///// </summary>
        ///// <param name="sqlDbType">sqlDbType</param>
        ///// <returns>C#数据类型</returns>
        //public static Type SqlTypeToCsharpType(SqlDbType sqlDbType)
        //{
        //    switch (sqlDbType)
        //    {
        //        case SqlDbType.BigInt:
        //            return typeof(Int64);
        //        case SqlDbType.Binary:
        //            return typeof(Object);
        //        case SqlDbType.Bit:
        //            return typeof(Boolean);
        //        case SqlDbType.Char:
        //            return typeof(String);
        //        case SqlDbType.DateTime:
        //            return typeof(DateTime);
        //        case SqlDbType.Decimal:
        //            return typeof(Decimal);
        //        case SqlDbType.Float:
        //            return typeof(Double);
        //        case SqlDbType.Image:
        //            return typeof(Object);
        //        case SqlDbType.Int:
        //            return typeof(Int32);
        //        case SqlDbType.Money:
        //            return typeof(Decimal);
        //        case SqlDbType.NChar:
        //            return typeof(String);
        //        case SqlDbType.NText:
        //            return typeof(String);
        //        case SqlDbType.NVarChar:
        //            return typeof(String);
        //        case SqlDbType.Real:
        //            return typeof(Single);
        //        case SqlDbType.SmallDateTime:
        //            return typeof(DateTime);
        //        case SqlDbType.SmallInt:
        //            return typeof(Int16);
        //        case SqlDbType.SmallMoney:
        //            return typeof(Decimal);
        //        case SqlDbType.Text:
        //            return typeof(String);
        //        case SqlDbType.Timestamp:
        //            return typeof(Object);
        //        case SqlDbType.TinyInt:
        //            return typeof(Byte);
        //        case SqlDbType.Udt:      //自定义的数据类型
        //            return typeof(Object);
        //        case SqlDbType.UniqueIdentifier:
        //            return typeof(Object);
        //        case SqlDbType.VarBinary:
        //            return typeof(Object);
        //        case SqlDbType.VarChar:
        //            return typeof(String);
        //        case SqlDbType.Variant:
        //            return typeof(Object);
        //        case SqlDbType.Xml:
        //            return typeof(Object);
        //        default:
        //            return null;
        //    }
        //}
        //#endregion

        //#region Type SqlTypeStringToCsharpType(string sqlTypeString)
        ///// <summary>
        ///// SQL Server中的数据类型字符串，转换为C#中的类型
        ///// </summary>
        ///// <param name="sqlTypeString">sql server数据类型字符串</param>
        ///// <returns>C#数据类型</returns>
        //public static Type SqlTypeStringToCsharpType(string sqlTypeString)
        //{
        //    SqlDbType dbTpe = SqlTypeStringToSqlDbType(sqlTypeString);
        //    return SqlTypeToCsharpType(dbTpe);
        //}
        //#endregion

        //#region string SqlTypeStringToCsharpTypeString(string sqlTypeString)
        ///// <summary>
        ///// 将sql server中的数据类型字符串，转化为C#中的类型的字符串
        ///// </summary>
        ///// <param name="sqlTypeString">sql server数据类型字符串</param>
        ///// <returns>C#数据类型字符串</returns>
        //public static string SqlTypeStringToCsharpTypeString(string sqlTypeString)
        //{
        //    Type type = SqlTypeStringToCsharpType(sqlTypeString);
        //    return type.Name;
        //}
        //#endregion

        //#region string SqlTypeStringToCsharpNullableTypeString(string sqlTypeString) 
        /////// <summary>
        /////// SQL Server中的数据类型字符串，转换为C#中的可空类型的字符串
        /////// </summary>
        /////// <param name="sqlTypeString"></param>
        /////// <returns></returns>
        ////public static string SqlTypeStringToCsharpNullableTypeString(string sqlTypeString)
        ////{
        ////    switch (sqlTypeString.ToLower())
        ////    {
        ////        case "int":
        ////        case "smallint":
        ////        case "bigint":
        ////            return "int?";
        ////        case "nvarchar":
        ////        case "nchar":
        ////        case "char":
        ////        case "varchar":
        ////        case "ntext":
        ////        case "text":
        ////            return "string";
        ////        case "money":
        ////        case "smallmoney":
        ////        case "decimal":
        ////        case "numeric":
        ////            return "decimal?";
        ////        case "float":
        ////            return "double?";
        ////        case "real":
        ////            return "single?";
        ////        case "datetime2":
        ////            return "DateTime2?"; ;
        ////        case "date":
        ////        case "smalldatetime":
        ////        case "datetime":
        ////            ;
        ////            return "DateTime?";
        ////        case "bit":
        ////            return "bool?";
        ////        case "tinyint":
        ////            return "byte?";
        ////        case "binary":
        ////        case "image":
        ////        case "rowversion":
        ////        case "timestamp":
        ////        case "varbinary":
        ////            return "byte[]";
        ////        case "uniqueidentifier":
        ////            return "Guid";
        ////        case "xml":
        ////            return "xml";
        ////        case "datetimeoffset":
        ////            return "datetimeoffset";
        ////        case "time":
        ////            return "TimeSpan";
        ////    }
        ////    MsgBox.ShowInformation("数据类型转化时含有未定义的类型:" + sqlTypeString);
        ////    return sqlTypeString;
        ////}
        //#endregion

        //#region string DataTypeToFullDataType
        ///// <summary>
        ///// 如：decimal 转化为decimal(18,4)
        ///// </summary>
        ///// <param name="dataType"></param>
        ///// <param name="maxLength"></param>
        ///// <param name="scale"></param>
        ///// <returns></returns>
        //public static string DataTypeToFullDataType(string dataType, int maxLength, int scale)
        //{
        //    switch (dataType)
        //    {
        //        case "varbinary":
        //        case "varchar":
        //        case "char":
        //        case "nchar":
        //        case "nvarchar":
        //            if (maxLength == -1)
        //            {
        //                return dataType + "(Max)";
        //            }
        //            return dataType + "(" + maxLength.ToString() + ")";

        //        case "numeric":
        //        case "decimal":
        //            return dataType + "(" + maxLength.ToString() + "," + scale.ToString() + ")";

        //        case "bit":
        //        case "int":
        //        case "datetime":
        //        case "bigint":
        //        case "tinyint":
        //        case "date":
        //        case "float":
        //        case "image":
        //        case "money":
        //        case "ntext":
        //        case "smalldatetime":
        //        case "text":
        //        case "timestamp":
        //        case "uniqueidentifier":
        //        case "smallint":
        //        case "smallmoney":
        //        case "real":
        //        case "time":
        //        case "xml":
        //        case "sql_variant":
        //        case "binary":
        //            return dataType;
        //        default:
        //            //MsgBox.ShowWarning("数据类型转化时含有未定义的类型:" + dataType);
        //            return dataType;
        //    }
        //}
        //#endregion

        //#region 把数据库中的类型转化为对应的DotNet类型(VB)
        ////public static string DBTypeToVBType(string dataType, out string valueType)
        ////{
        ////    switch (dataType.ToLower())
        ////    {
        ////        case "int":
        ////        case "smallint":
        ////        case "bigint":
        ////            valueType = "0";
        ////            return "Int64";
        ////        case "nvarchar":
        ////        case "nchar":
        ////        case "char":
        ////        case "varchar":
        ////        case "ntext":
        ////        case "text":
        ////            valueType = "Nothing";
        ////            return "String";
        ////        case "money":
        ////        case "smallmoney":
        ////        case "decimal":
        ////        case "numeric":
        ////            valueType = "0";
        ////            return "Decimal";
        ////        case "float":
        ////            valueType = "0";
        ////            return "Double";
        ////        case "real":
        ////            valueType = "0";
        ////            return "Single";
        ////        case "datetime2":
        ////            valueType = "Nothing";
        ////            return "DateTime2"; ;
        ////        case "date":
        ////        case "smalldatetime":
        ////        case "datetime":
        ////            valueType = "Nothing";
        ////            return "DateTime";
        ////        case "bit":
        ////            valueType = "False";
        ////            return "Boolean";
        ////        case "tinyint":
        ////            valueType = "Nothing";
        ////            return "Byte";
        ////        case "binary":
        ////        case "image":
        ////        case "rowversion":
        ////        case "timestamp":
        ////        case "varbinary":
        ////            valueType = "Nothing";
        ////            return "Byte()";
        ////        case "uniqueidentifier":
        ////            valueType = "Nothing";
        ////            return "Guid";
        ////        case "xml":
        ////            valueType = "Nothing";
        ////            return "xml";
        ////        case "datetimeoffset":
        ////            valueType = "Nothing";
        ////            return "DateTimeOffset";
        ////        case "time":
        ////            valueType = "Nothing";
        ////            return "TimeSpan";
        ////    }
        ////    MsgBox.ShowInformation("数据类型转化时含有未定义的类型:" + dataType);
        ////    valueType = "Nothing";
        ////    return dataType;
        ////}
        //#endregion

        //#region 生成存储过程时的数据类型转化
        /////// <summary>
        /////// 
        /////// </summary>
        /////// <param name="drColumnInfo">dr必须为SqlTableColumnInfo 中的行</param>
        /////// <returns></returns>
        ////public static string DBTypeToProcType(DataRow drColumnInfo)
        ////{
        ////    string convertedType = "";
        ////    if (drColumnInfo == null)
        ////    {
        ////        MsgBox.ShowInformation("传入的列不能为空");
        ////    }
        ////    string dbType = drColumnInfo["ColumnType"].ToString().ToLower();
        ////    string dataLength = drColumnInfo["ColumnLength"].ToString();
        ////    string digitScale = drColumnInfo["Scale"].ToString(); //小數位位數

        ////    switch (dbType)
        ////    {
        ////        case "varbinary":
        ////        case "varchar":
        ////        case "char":
        ////        case "nchar":
        ////        case "nvarchar":
        ////            if (dataLength == "-1")
        ////            {
        ////                dataLength = "Max";
        ////            }
        ////            convertedType = dbType + "(" + dataLength + ")";
        ////            return convertedType;

        ////        case "numeric":
        ////        case "decimal":
        ////            convertedType = dbType + "(" + dataLength + "," + digitScale + ")";
        ////            return convertedType;

        ////        case "bit":
        ////        case "int":
        ////        case "datetime":
        ////        case "bigint":
        ////        case "tinyint":
        ////        case "date":
        ////        case "float":
        ////        case "image":
        ////        case "money":
        ////        case "ntext":
        ////        case "smalldatetime":
        ////        case "text":
        ////        case "timestamp":
        ////        case "uniqueidentifier":
        ////        case "smallint":
        ////        case "smallmoney":
        ////        case "real":
        ////        case "time":
        ////        case "xml":
        ////            return convertedType = dbType;
        ////        default:
        ////            MsgBox.ShowWarning("数据类型转化时含有未定义的类型:" + dbType);
        ////            return convertedType = dbType;
        ////    }
        ////}
        //#endregion
        #endregion
    }
}
