using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace CommonExpand
{
    /// <summary>
    /// 常用正则表达式验证
    /// </summary>
    public static class RegexCommon
    {
        #region 正则校验拓展
        private static Regex reg { get; set; }
        /// <summary>
        /// 自定义正则表达式规则
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则</param>
        /// <returns></returns>
        public static bool IsRegex(this string str, Regex regex) => regex.IsMatch(str);
        /// <summary>
        /// 自定义正则表达式规则
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则</param>
        /// <returns></returns>
        public static bool IsRegex(this string str, string regex) => new Regex(regex).IsMatch(str);
        /// <summary>
        /// 正则判断是否是正整数,不含0（例：1,23,456）
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsPositiveIntegerReg(this string str) => new Regex(@"^[1-9]\d*$", RegexOptions.IgnoreCase).IsMatch(str);
        /// <summary>
        /// 正则判断是否是纯数字（例：123456）
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsPureNumberReg(this string str)
        {
            reg = new Regex(@"^[0 - 9]*$");
            if (reg.IsMatch(str)) return true;
            return false;
        }
        /// <summary>
        /// 正则判断是否是数字,含正负整数、正负小数（例：-123.456）
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsNumberReg(this string str)
        {
            reg = new Regex(@"^[+-]?\d+[.]?\d*$");
            if (reg.IsMatch(str)) return true;
            return false;
        }
        /// <summary>
        /// 是否是小数,可带+-
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsDecimalReg(this string str)
        {
            Match m = new Regex("^[+-]?[0-9]+[.]?[0-9]+$").Match(str);
            return m.Success;
        }
        /// <summary>
        /// 正则判断是否是日期（例：2000-01-01或2000/01/01或2000.01.01）
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则</param>
        /// <returns></returns>
        public static bool IsDateReg(this string str, Regex regex = null)
        {
            if (regex == null) reg = new Regex(@"^\d{4}(-|/|.)\d{1,2}(-|/|.)\d{1,2}$");
            if (reg.IsMatch(str)) return true;
            return false;
        }
        /// <summary>
        /// 正则判断是否是11位手机号码
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则,默认:^(13[0-9]|14[5|7]|15[0|1|2|3|5|6|7|8|9]|17[0|1|2|3|5|6|7|8|9]|18[0|1|2|3|5|6|7|8|9]|19[0|1|2|3|5|6|7|8|9])\d{8}$</param>
        /// <returns></returns>
        public static bool IsPhoneReg(this string str, Regex regex = null)
        {
            if (regex == null) reg = new Regex(@"^(13[0-9]|14[5|7]|15[0|1|2|3|5|6|7|8|9]|17[0|1|2|3|5|6|7|8|9]|18[0|1|2|3|5|6|7|8|9]|19[0|1|2|3|5|6|7|8|9])\d{8}$");
            if (reg.IsMatch(str)) return true;
            return false;
        }
        /// <summary>
        /// 正则判断是否是邮箱
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则,默认:^[\\w-]+@[\\w-]+\\.(com|net|org|edu|mil|tv|biz|info)$</param>
        /// <returns></returns>
        public static bool IsEmailReg(this string str, Regex regex = null)
        {
            Regex RegEmail = new Regex("^[\\w-]+@[\\w-]+\\.(com|net|org|edu|mil|tv|biz|info)$");//w 英文字母或数字的字符串，和 [a-zA-Z0-9] 语法一样 
            if (regex != null) RegEmail = regex;
            Match m = RegEmail.Match(str);
            return m.Success;
        }
        /// <summary>
        /// 正则判断是否是IP地址
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则,默认:^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$</param>
        /// <returns></returns>
        public static bool IsIPAddressReg(this string str, Regex regex = null)
        {
            Regex RegEmail = new Regex(@"^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$");
            if (regex != null) RegEmail = regex;
            Match m = RegEmail.Match(str);
            return m.Success;
        }
        /// <summary>
        /// 正则判断密码有效性,A-Z,a-z,0-9(6-16位)
        /// </summary>
        /// <param name="password">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则,默认:^[A-Za-z_0-9]{6,16}$</param>
        /// <returns></returns>
        public static bool IsValidPasswordReg(this string password, Regex regex = null)
        {
            if (regex != null) return regex.IsMatch(password);
            return Regex.IsMatch(password, @"^[A-Za-z_0-9]{6,16}$");
        }
        /// <summary>
        /// 正则判断是否是全中文
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则,默认:[\u4e00-\u9fa5]</param>
        /// <returns></returns>
        public static bool IsChineseReg(this string str, Regex regex = null)
        {
            Regex RegEmail = new Regex("[\u4e00-\u9fa5]");
            if (regex != null) RegEmail = regex;
            Match m = RegEmail.Match(str);
            return m.Success;
        }
        /// <summary>
        /// 正则判断Url有效性
        /// </summary>
        /// <param name="url">需要判定的字符串</param>
        /// <param name="regex">自定义正则验证规则,默认:^(http|https|ftp)\://[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,3}(:[a-zA-Z0-9]*)?/?([a-zA-Z0-9\-\._\?\,\'/\\\+&%\$#\=~])*[^\.\,\)\(\s]$</param>
        /// <returns></returns>
        static public bool IsValidUrlReg(this string url, Regex regex = null)
        {
            if (regex != null) return regex.IsMatch(url);
            return Regex.IsMatch(url, @"^(http|https|ftp)\://[a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,3}(:[a-zA-Z0-9]*)?/?([a-zA-Z0-9\-\._\?\,\'/\\\+&%\$#\=~])*[^\.\,\)\(\s]$");
        }

        #endregion
    }
}