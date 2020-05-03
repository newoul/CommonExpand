using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CommonExpand
{
    /// <summary>
    /// 数据获取
    /// </summary>
    public static partial class DataAcquisition
    {
        #region 枚举

        #region 枚举成员转成dictionary类型
        /// <summary>
        /// 转成dictionary类型
        /// </summary>
        /// <param name="enumType"></param>
        /// <returns></returns>
        public static Dictionary<int, string> EnumToDictionary(this Type enumType)
        {
            Dictionary<int, string> dictionary = new Dictionary<int, string>();
            Type typeDescription = typeof(DescriptionAttribute);
            FieldInfo[] fields = enumType.GetFields();
            int sValue = 0;
            string sText = string.Empty;
            foreach (FieldInfo field in fields)
            {
                if (field.FieldType.IsEnum)
                {
                    sValue = ((int)enumType.InvokeMember(field.Name, BindingFlags.GetField, null, null, null));
                    object[] arr = field.GetCustomAttributes(typeDescription, true);
                    if (arr.Length > 0)
                    {
                        DescriptionAttribute da = (DescriptionAttribute)arr[0];
                        sText = da.Description;
                    }
                    else
                    {
                        sText = field.Name;
                    }
                    dictionary.Add(sValue, sText);
                }
            }
            return dictionary;
        }
        /// <summary>
        /// 枚举成员转成键值对Json字符串
        /// </summary>
        /// <param name="enumType"></param>
        /// <returns></returns>
        public static string EnumToDictionaryString(this Type enumType)
        {
            List<KeyValuePair<int, string>> dictionaryList = EnumToDictionary(enumType).ToList();
            var sJson = JsonConvert.SerializeObject(dictionaryList);
            return sJson;
        }
        #endregion

        #region 获取枚举的描述
        /// <summary>
        /// 获取枚举值对应的描述
        /// </summary>
        /// <param name="enumType"></param>
        /// <returns></returns>
        public static string GetDescription(this System.Enum enumType)
        {
            FieldInfo EnumInfo = enumType.GetType().GetField(enumType.ToString());
            if (EnumInfo != null)
            {
                DescriptionAttribute[] EnumAttributes = (DescriptionAttribute[])EnumInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (EnumAttributes.Length > 0)
                {
                    return EnumAttributes[0].Description;
                }
            }
            return enumType.ToString();
        }
        #endregion

        #region 根据值获取枚举的描述
        /// <summary>
        /// 获取枚举的<see cref="System.ComponentModel.DescriptionAttribute"/>的描述
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string GetDescriptionByEnum<T>(this object obj)
        {
            var tEnum = System.Enum.Parse(typeof(T), obj.ParseToString()) as System.Enum;
            var description = tEnum.GetDescription();
            return description;
        }
        #endregion

        #endregion

        #region String 身份证号识别拓展

        /// <summary>
        /// 根据身份证号获取出生日期，返回yyyy-MM-dd,未匹配到或转化日期失败返回<see cref="string.Empty"/>
        /// </summary>
        /// <param name="idcardStr"></param>
        /// <returns></returns>
        public static string GetIDCardBirthday(this string idcardStr) {
            string _birthday = string.Empty;
            if (idcardStr.IsIDCard()) {
                if (idcardStr.Length == 18)//处理18位的身份证号码从号码中得到生日和性别代码
                {
                    _birthday = idcardStr.Substring(6, 4) + "-" + idcardStr.Substring(10, 2) + "-" + idcardStr.Substring(12, 2);
                    if(!DateTime.TryParse(_birthday, out _))return string.Empty;
                }
                if (idcardStr.Length == 15)
                {
                    _birthday = "19" + idcardStr.Substring(6, 2) + "-" + idcardStr.Substring(8, 2) + "-" + idcardStr.Substring(10, 2);
                    if (!DateTime.TryParse(_birthday, out _)) return string.Empty;
                }
            }
            return _birthday;
        }

        /// <summary>
        /// 根据身份证号获取性别，返回“男”或“女”,无法识别返回“未知”
        /// </summary>
        /// <param name="idcardStr"></param>
        /// <returns></returns>
        public static string GetIDCardGender(this string idcardStr)
        {
            string _sex = "未知";
            int _gender = 0;
            if (idcardStr.IsIDCard())
            {
                if (idcardStr.Length == 18)//处理18位的身份证号码从号码中得到生日和性别代码
                {
                    _sex = idcardStr.Substring(14, 3);
                    if(int.TryParse(_sex,out _gender)) return _gender% 2 == 0 ? "女" : "男";
                }
                if (idcardStr.Length == 15)
                {
                    _sex = idcardStr.Substring(12, 3);
                    if (int.TryParse(_sex, out _gender)) return _gender % 2 == 0 ? "女" : "男";
                }
            }
            return _sex;
        }
        /// <summary>
        /// 根据身份证号获取性别，返回“1=男”或“2=女”,无法识别返回“0=未知”
        /// </summary>
        /// <param name="idcardStr"></param>
        /// <returns></returns>
        public static int GetIDCardGenderInt(this string idcardStr)
        {
            string _sex = string.Empty;
            int _gender = 0;
            if (idcardStr.IsIDCard())
            {
                if (idcardStr.Length == 18)//处理18位的身份证号码从号码中得到生日和性别代码
                {
                    _sex = idcardStr.Substring(14, 3);
                    if (int.TryParse(_sex, out _gender)) return _gender % 2 == 0 ? 2 : 1;
                }
                if (idcardStr.Length == 15)
                {
                    _sex = idcardStr.Substring(12, 3);
                    if (int.TryParse(_sex, out _gender)) return _gender % 2 == 0 ? 2 : 1;
                }
            }
            return _gender;
        }

        #endregion

        #region DateTime类型拓展

        /// <summary>
        /// 转为日期时间,如果为Null返回<see cref="String.Empty"/>
        /// </summary>
        /// <param name="dt">当前要转的<see cref="DateTime"/></param>
        /// <param name="format">定义转化格式</param>
        /// <returns></returns>
        public static string ToFormatDate(this DateTime? dt, string format = "yyyy-MM-dd")
        {
            if (dt.IsNotNull() && format.IsNotNullAndEmpty()) return dt.Value.ToString(format);
            return string.Empty;
        }
        /// <summary>
        /// 转为日期时间
        /// </summary>
        /// <param name="dt">当前要转的<see cref="DateTime"/></param>
        /// <param name="format">定义转化格式</param>
        /// <returns></returns>
        public static string ToFormatDate(this DateTime dt, string format = "yyyy-MM-dd")=>dt.ToString(format);
        /// <summary>
        /// 转为日期时间,如果为Null返回<see cref="String.Empty"/>
        /// </summary>
        /// <param name="dt">当前要转的<see cref="DateTime"/></param>
        /// <param name="format">定义转化格式</param>
        /// <returns></returns>
        public static string ToFormatDateTime(this DateTime? dt, string format = "yyyy-MM-dd HH:mm:ss") => dt.ToFormatDate(format);
        /// <summary>
        /// 转为日期时间
        /// </summary>
        /// <param name="dt">当前要转的<see cref="DateTime"/></param>
        /// <param name="format">定义转化格式</param>
        /// <returns></returns>
        public static string ToFormatDateTime(this DateTime dt, string format = "yyyy-MM-dd HH:mm:ss") => dt.ToFormatDate(format);
        /// <summary>
        /// 获取当前日期的最大时间 23:59:59
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime GetDayMaxTime(DateTime dt) => new DateTime(dt.Year, dt.Month, dt.Day, 23, 59, 59);
        /// <summary>
        /// 获取当前日期的下一天。 NextDay：yyyy-MM-dd 00:00:00
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime GetNextDate(DateTime dt) => new DateTime(dt.Year, dt.Month, dt.Day).AddDays(1);
        /// <summary>
        /// 获取当前日期的前一天。 UpDay：yyyy-MM-dd 00:00:00
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime GetUpDate(DateTime dt) => new DateTime(dt.Year, dt.Month, dt.Day).AddDays(-1);
        /// <summary>
        /// 获取当前月第一天,假如今天是2000-01-15，这当月最后一天为2000-01-01 00:00:00
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime GetMonthFirstDay(this DateTime dt) => new DateTime(dt.Year, dt.Month, 1);
        /// <summary>
        /// 获取当前月最后一天, 假如今天是2000-01-15，这当月最后一天为2000-01-31 23:59:59
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DateTime GetMonthLastDay(this DateTime dt) => dt.GetMonthFirstDay().AddMonths(1).AddSeconds(-1);
        /// <summary>
        /// 获取当前月第一天, 字符串格式。假如今天是2000-01-15，这当月最后一天为2000-01-01 00:00:00
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string GetMonthFirstDayStr(this DateTime dt) => new DateTime(dt.Year, dt.Month, 1).ToShortDateString();
        /// <summary>
        /// 获取当前月最后一天, 字符串格式。假如今天是2000-01-15，这当月最后一天为2000-01-31 23:59:59
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string GetMonthLastDayStr(this DateTime dt) => dt.GetMonthLastDay().ToShortDateString();
        /// <summary>
        /// 获取当前日期是星期几，返回（1,2,3,4,5,6,7）
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static int GetWeek(this DateTime dt) => dt.DayOfWeek==DayOfWeek.Sunday?7:(int)dt.DayOfWeek;

        #endregion

        #region 异常获取

        /// <summary>
        ///获取最底层异常
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        public static Exception GetOriginalException(this Exception ex)
        {
            if (ex.InnerException == null) return ex;

            return ex.InnerException.GetOriginalException();
        }

        /// <summary>
        /// 获取最底层真实错误信息
        /// </summary>
        /// <param name="exception"></param>
        /// <returns></returns>
        public static string GetOriginalRealError(this Exception exception)
        {
            return exception.GetOriginalException().Message;
        }
        /// <summary>
        /// 获取最底层堆栈错误信息
        /// </summary>
        /// <param name="exception"></param>
        /// <returns></returns>
        public static string GetOriginalStackTrace(this Exception exception)
        {
            return exception.GetOriginalException().StackTrace;
        }
        /// <summary>
        /// 获取真实错误信息(6层)
        /// </summary>
        /// <param name="exception"></param>
        /// <returns></returns>
        public static string GetRealError(this Exception exception)
        {
            bool flag = exception == null;
            string result;
            if (flag)
            {
                result = string.Empty;
            }
            else
            {
                bool flag2 = exception.InnerException != null;
                if (flag2)
                {
                    bool flag3 = exception.InnerException.InnerException != null;
                    if (flag3)
                    {
                        bool flag4 = exception.InnerException.InnerException.InnerException != null;
                        if (flag4)
                        {
                            bool flag5 = exception.InnerException.InnerException.InnerException.InnerException != null;
                            if (flag5)
                            {
                                bool flag6 = exception.InnerException.InnerException.InnerException.InnerException.InnerException != null;
                                if (flag6)
                                {
                                    bool flag7 = exception.InnerException.InnerException.InnerException.InnerException.InnerException.InnerException != null;
                                    if (flag7)
                                    {
                                        result = exception.InnerException.InnerException.InnerException.InnerException.InnerException.InnerException.Message;
                                    }
                                    else
                                    {
                                        result = exception.InnerException.InnerException.InnerException.InnerException.InnerException.Message;
                                    }
                                }
                                else
                                {
                                    result = exception.InnerException.InnerException.InnerException.InnerException.Message;
                                }
                            }
                            else
                            {
                                result = exception.InnerException.InnerException.InnerException.Message;
                            }
                        }
                        else
                        {
                            result = exception.InnerException.InnerException.Message;
                        }
                    }
                    else
                    {
                        result = exception.InnerException.Message;
                    }
                }
                else
                {
                    result = exception.Message;
                }
            }
            return result;
        }
        /// <summary>
        /// 获取堆栈错误信息(6层)
        /// </summary>
        /// <param name="exception"></param>
        /// <returns></returns>
        public static string GetStackTrace(this Exception exception)
        {
            bool flag = exception == null;
            string result;
            if (flag)
            {
                result = string.Empty;
            }
            else
            {
                bool flag2 = exception.InnerException != null;
                if (flag2)
                {
                    bool flag3 = exception.InnerException.InnerException != null;
                    if (flag3)
                    {
                        bool flag4 = exception.InnerException.InnerException.InnerException != null;
                        if (flag4)
                        {
                            bool flag5 = exception.InnerException.InnerException.InnerException.InnerException != null;
                            if (flag5)
                            {
                                bool flag6 = exception.InnerException.InnerException.InnerException.InnerException.InnerException != null;
                                if (flag6)
                                {
                                    bool flag7 = exception.InnerException.InnerException.InnerException.InnerException.InnerException.InnerException != null;
                                    if (flag7)
                                    {
                                        result = exception.InnerException.InnerException.InnerException.InnerException.InnerException.InnerException.StackTrace;
                                    }
                                    else
                                    {
                                        result = exception.InnerException.InnerException.InnerException.InnerException.InnerException.StackTrace;
                                    }
                                }
                                else
                                {
                                    result = exception.InnerException.InnerException.InnerException.InnerException.StackTrace;
                                }
                            }
                            else
                            {
                                result = exception.InnerException.InnerException.InnerException.StackTrace;
                            }
                        }
                        else
                        {
                            result = exception.InnerException.InnerException.StackTrace;
                        }
                    }
                    else
                    {
                        result = exception.InnerException.StackTrace;
                    }
                }
                else
                {
                    result = exception.StackTrace;
                }
            }
            return result;
        }

        #endregion
    }
}
