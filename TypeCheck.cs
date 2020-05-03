using System;
using System.Collections.Generic;
using System.Text;

namespace CommonExpand
{
    /// <summary>
    /// 类型校验
    /// </summary>
    public static partial class TypeCheck
    {
        #region 是否为某类型  返回Bool

        /// <summary>
        /// 判断是否为null或<see cref="string.Empty"/>空值
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsNullOrEmpty(this string str) => string.IsNullOrEmpty(str);

        /// <summary>
        /// 不是null和<see cref="string.Empty"/>则为true,否则为false
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsNotNullAndEmpty(this string str) => !string.IsNullOrEmpty(str);
        /// <summary>
        /// 是Null则返回true
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsNull(this string str) => str == null;
        /// <summary>
        ///  不是Null则返回true
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsNotNull(this string str) => str != null;
        /// <summary>
        /// 是Null或0
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsNullOrZero(this object value) => (value == null || value.ParseToString().Trim() == "0");

        /// <summary>
        /// 是Null则返回true
        /// </summary>
        /// <param name="obj">校验值</param>
        /// <returns></returns>
        public static bool IsNull(this object obj) => (obj == null || obj is DBNull);
        /// <summary>
        ///  不是Null则返回true
        /// </summary>
        /// <param name="obj">校验值</param>
        /// <returns></returns>
        public static bool IsNotNull(this object obj) => obj != null && !(obj is DBNull);
        /// <summary>
        /// 判断是否是<see cref="int"/>类型
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsInt(this string str)
        {
            int _val = 0;
            return int.TryParse(str, out _val);
        }

        /// <summary>
        /// 判断是否是<see cref="decimal"/>类型
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsDecimal(this string str)
        {
            decimal _val = 0;
            return decimal.TryParse(str, out _val);
        }

        /// <summary>
        /// 判断是否是<see cref="double"/>类型
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsDouble(this string str)
        {
            double _val = 0;
            return double.TryParse(str, out _val);
        }

        /// <summary>
        /// 判断是否是<see cref="DateTime"/>类型
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsDateTime(this string str)
        {
            DateTime _date = new DateTime();
            return DateTime.TryParse(str, out _date);
        }

        /// <summary>
        /// 判断是否是<see cref="Guid"/>类型
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsGuid(this string str)
        {
            if (string.IsNullOrEmpty(str)) return false;
            try
            {
                Guid _t = new Guid(str);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 验证身份证合理性，含18位和15位身份证验证
        /// </summary>
        /// <param name="str">需要判定的字符串</param>
        /// <returns></returns>
        public static bool IsIDCard(this string str)
        {
            if (str.IsNotNullAndEmpty() && str.Length == 18)
            {
                bool check = CheckIDCard18(str);
                return check;
            }
            else if (str.IsNotNullAndEmpty() && str.Length == 15)
            {
                bool check = CheckIDCard15(str);
                return check;
            }
            else
            {
                return false;
            }
        }
        /// <summary>  
        /// 18位身份证号码验证  
        /// </summary>  
        public static bool CheckIDCard18(this string idNumber)
        {
            long n = 0;
            if (long.TryParse(idNumber.Remove(17), out n) == false
                || n < Math.Pow(10, 16) || long.TryParse(idNumber.Replace('x', '0').Replace('X', '0'), out n) == false)
            {
                return false;//数字验证  
            }
            string address = "11x22x35x44x53x12x23x36x45x54x13x31x37x46x61x14x32x41x50x62x15x33x42x51x63x21x34x43x52x64x65x71x81x82x91";
            if (address.IndexOf(idNumber.Remove(2)) == -1)
            {
                return false;//省份验证  
            }
            string birth = idNumber.Substring(6, 8).Insert(6, "-").Insert(4, "-");
            DateTime time = new DateTime();
            if (DateTime.TryParse(birth, out time) == false)
            {
                return false;//生日验证  
            }
            string[] arrVarifyCode = ("1,0,x,9,8,7,6,5,4,3,2").Split(',');
            string[] Wi = ("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2").Split(',');
            char[] Ai = idNumber.Remove(17).ToCharArray();
            int sum = 0;
            for (int i = 0; i < 17; i++)
            {
                sum += int.Parse(Wi[i]) * int.Parse(Ai[i].ToString());
            }
            int y = -1;
            Math.DivRem(sum, 11, out y);
            if (arrVarifyCode[y] != idNumber.Substring(17, 1).ToLower())
            {
                return false;//校验码验证  
            }
            return true;//符合GB11643-1999标准  
        }

        /// <summary>  
        /// 16位身份证号码验证  
        /// </summary>  
        public static bool CheckIDCard15(this string idNumber)
        {
            long n = 0;
            if (long.TryParse(idNumber, out n) == false || n < Math.Pow(10, 14))
            {
                return false;//数字验证  
            }
            string address = "11x22x35x44x53x12x23x36x45x54x13x31x37x46x61x14x32x41x50x62x15x33x42x51x63x21x34x43x52x64x65x71x81x82x91";
            if (address.IndexOf(idNumber.Remove(2)) == -1)
            {
                return false;//省份验证  
            }
            string birth = idNumber.Substring(6, 6).Insert(4, "-").Insert(2, "-");
            DateTime time = new DateTime();
            if (DateTime.TryParse(birth, out time) == false)
            {
                return false;//生日验证  
            }
            return true;
        }

        /// <summary>
        /// 是否是闰年，值为Null也会返回false
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool IsLeapYear(this DateTime? dt)
        {
            if (dt.IsNotNull()) return dt.Value.IsLeapYear();
            return false;
        }
        /// <summary>
        /// 是否是闰年
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static bool IsLeapYear(this DateTime dt)=> (dt.Year % 400 == 0 || dt.Year % 4 == 0 && dt.Year % 100 != 0);

        #endregion
    }
}
