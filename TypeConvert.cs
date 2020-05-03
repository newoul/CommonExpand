using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace CommonExpand
{
    /// <summary>
    /// 数据转换
    /// </summary>
    public static partial class TypeConvert
    {
        #region String转各种类型

        #region 转换为long
        /// <summary>
        /// 将object转换为long，若转换失败，则返回0。不抛出异常。  
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static long ParseToLong(this string obj)
        {
            try
            {
                return long.Parse(obj.ToString());
            }
            catch
            {
                return 0L;
            }
        }

        /// <summary>
        /// 将object转换为long，若转换失败，则返回指定值。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static long ParseToLong(this string str, long defaultValue)
        {
            try
            {
                return long.Parse(str);
            }
            catch
            {
                return defaultValue;
            }
        }
        #endregion

        #region 转换为int
        /// <summary>
        /// 将object转换为int，若转换失败，则返回0。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static int ParseToInt(this string str)
        {
            try
            {
                return Convert.ToInt32(str);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 将object转换为int，若转换失败，则返回指定值。不抛出异常。 
        /// null返回默认值
        /// </summary>
        /// <param name="str"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        public static int ParseToInt(this string str, int defaultValue)
        {
            if (str == null)
            {
                return defaultValue;
            }
            try
            {
                return Convert.ToInt32(str);
            }
            catch
            {
                return defaultValue;
            }
        }
        #endregion

        #region 转换为short
        /// <summary>
        /// 将object转换为short，若转换失败，则返回0。不抛出异常。  
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static short ParseToShort(this string obj)
        {
            try
            {
                return short.Parse(obj.ToString());
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 将object转换为short，若转换失败，则返回指定值。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <param name="defaultValue">默认值</param>
        /// <returns></returns>
        public static short ParseToShort(this string str, short defaultValue)
        {
            try
            {
                return short.Parse(str.ToString());
            }
            catch
            {
                return defaultValue;
            }
        }
        #endregion

        #region 转换为demical
        /// <summary>
        /// 将object转换为demical，若转换失败，则返回指定值。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <param name="defaultValue">默认值</param>
        /// <returns></returns>
        public static decimal ParseToDecimal(this string str, decimal defaultValue)
        {
            try
            {
                return decimal.Parse(str.ToString());
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// 将object转换为demical，若转换失败，则返回0。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static decimal ParseToDecimal(this string str)
        {
            try
            {
                return decimal.Parse(str.ToString());
            }
            catch
            {
                return 0;
            }
        }
        #endregion

        #region 转化为bool
        /// <summary>
        /// 将object转换为bool，若转换失败，则返回false。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool ParseToBool(this string str)
        {
            try
            {
                return bool.Parse(str.ToString());
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 将object转换为bool，若转换失败，则返回指定值。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <param name="defaultValue">默认值</param>
        /// <returns></returns>
        public static bool ParseToBool(this string str, bool defaultValue)
        {
            try
            {
                return bool.Parse(str.ToString());
            }
            catch
            {
                return defaultValue;
            }
        }
        #endregion

        #region 转换为float
        /// <summary>
        /// 将object转换为float，若转换失败，则返回0。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static float ParseToFloat(this string str)
        {
            try
            {
                return float.Parse(str.ToString());
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 将object转换为float，若转换失败，则返回指定值。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <param name="defaultValue">默认值</param>
        /// <returns></returns>
        public static float ParseToFloat(this string str, float defaultValue)
        {
            try
            {
                return float.Parse(str.ToString());
            }
            catch
            {
                return defaultValue;
            }
        }
        #endregion

        #region 转换为Guid
        /// <summary>
        /// 将string转换为Guid，若转换失败，则返回Guid.Empty。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static Guid ParseToGuid(this string str)
        {
            try
            {
                return new Guid(str);
            }
            catch
            {
                return Guid.Empty;
            }
        }
        #endregion

        #region 转换为DateTime
        /// <summary>
        /// 将string转换为DateTime，若转换失败，则返回日期最小值。不抛出异常。 
        /// 常见的有：yyyy(-/.)MM(-/.)dd；yyyy(-/.)MM(-/.)dd HH:mm:ss；yyyyMMdd；yyyyMMddHHmmss
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static DateTime ParseToDateTime(this string str)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(str))
                {
                    return DateTime.MinValue;
                }
                if (str.Contains("-") || str.Contains("/") || str.Contains("."))
                {
                    return DateTime.Parse(str);
                }
                else
                {
                    int length = str.Length;
                    switch (length)
                    {
                        case 4:
                            return DateTime.ParseExact(str, "yyyy", System.Globalization.CultureInfo.CurrentCulture);
                        case 6:
                            return DateTime.ParseExact(str, "yyyyMM", System.Globalization.CultureInfo.CurrentCulture);
                        case 8:
                            return DateTime.ParseExact(str, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);
                        case 10:
                            return DateTime.ParseExact(str, "yyyyMMddHH", System.Globalization.CultureInfo.CurrentCulture);
                        case 12:
                            return DateTime.ParseExact(str, "yyyyMMddHHmm", System.Globalization.CultureInfo.CurrentCulture);
                        case 14:
                            return DateTime.ParseExact(str, "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
                        default:
                            return DateTime.ParseExact(str, "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
                    }
                }
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// 将string转换为DateTime，若转换失败，则返回默认值。  
        /// 常见的有：yyyy(-/.)MM(-/.)dd；yyyy(-/.)MM(-/.)dd HH:mm:ss；yyyyMMdd；yyyyMMddHHmmss
        /// </summary>
        /// <param name="str"></param>
        /// <param name="defaultValue">默认值</param>
        /// <returns></returns>
        public static DateTime ParseToDateTime(this string str, DateTime? defaultValue)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(str))
                {
                    return defaultValue.GetValueOrDefault();
                }
                if (str.Contains("-") || str.Contains("/") || str.Contains("."))
                {
                    return DateTime.Parse(str);
                }
                else
                {
                    int length = str.Length;
                    switch (length)
                    {
                        case 4:
                            return DateTime.ParseExact(str, "yyyy", System.Globalization.CultureInfo.CurrentCulture);
                        case 6:
                            return DateTime.ParseExact(str, "yyyyMM", System.Globalization.CultureInfo.CurrentCulture);
                        case 8:
                            return DateTime.ParseExact(str, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);
                        case 10:
                            return DateTime.ParseExact(str, "yyyyMMddHH", System.Globalization.CultureInfo.CurrentCulture);
                        case 12:
                            return DateTime.ParseExact(str, "yyyyMMddHHmm", System.Globalization.CultureInfo.CurrentCulture);
                        case 14:
                            return DateTime.ParseExact(str, "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
                        default:
                            return DateTime.ParseExact(str, "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
                    }
                }
            }
            catch
            {
                return defaultValue.GetValueOrDefault();
            }
        }
        #endregion

        #region 转换为double
        /// <summary>
        /// 将object转换为double，若转换失败，则返回0。不抛出异常。  
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static double ParseToDouble(this string obj)
        {
            try
            {
                return double.Parse(obj.ToString());
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 将object转换为double，若转换失败，则返回指定值。不抛出异常。  
        /// </summary>
        /// <param name="str"></param>
        /// <param name="defaultValue">默认值</param>
        /// <returns></returns>
        public static double ParseToDouble(this string str, double defaultValue)
        {
            try
            {
                return double.Parse(str.ToString());
            }
            catch
            {
                return defaultValue;
            }
        }
        #endregion

        #region 各类型转换为string
        /// <summary>
        /// 将object转换为string，若转换失败，则返回""。不抛出异常。  
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ParseToString(this object obj)
        {
            try
            {
                if (obj == null)
                {
                    return string.Empty;
                }
                else
                {
                    return obj.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
        }
        /// <summary>
        /// 数组转字符串
        /// </summary>
        /// <typeparam name="T">数组</typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ParseToStrings<T>(this IEnumerable<T> obj)
        {
            try
            {
                var list = obj as IEnumerable<T>;
                if (list != null)
                {
                    return string.Join(",", list);
                }
                else
                {
                    return obj.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }

        }
        #endregion

        //public static void test() {
        //    var ss = new int[]{ 1,2,3};
        //  var bb=  ss.ParseToStrings();
        //}
        #endregion

        #region 强制转换类型
        /// <summary>
        /// 强制转换类型
        /// </summary>
        /// <typeparam name="TResult">强制转换的类型</typeparam>
        /// <param name="source">当前数据源</param>
        /// <returns></returns>
        public static IEnumerable<TResult> CastSuper<TResult>(this IEnumerable source)
        {
            foreach (object item in source)
            {
                yield return (TResult)Convert.ChangeType(item, typeof(TResult));
            }
        }
        #endregion
    }
}
