using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonExpand
{
    /// <summary>
    /// 印章帮助类，.net core没有System.Drawing，替代方案是使用第三方的依赖，这里推荐使用Nuget的System.Drawing.Common。
    /// </summary>
    public class SealHelper
    {
        /// <summary>
        /// 单位印章初始化_自定义
        /// </summary>
        /// <returns></returns>
        public static MechanismSealHelper Mechanism()
        {
            return new MechanismSealHelper();
        }
        /// <summary>
        /// 个人印章初始化_自定义
        /// </summary>
        /// <returns></returns>
        public static PersonalSealHelper Personal()
        {
            return new PersonalSealHelper();
        }

        /// <summary>
        /// 单位印章转<see cref="byte"/>流
        /// </summary>
        /// <param name="companyName">公司名称</param>
        /// <param name="centerText">中间文字</param>
        /// <param name="bottomText">下弦文字</param>
        /// <param name="borderShow">是否显示圆形边框</param>
        /// <param name="starShow">是否显示星星</param>
        /// <returns></returns>
        public static byte[] MechanismToByte(string companyName, string centerText = "", string bottomText = "", bool borderShow = true, bool starShow = true)
        {
            using (var helper = SealHelper.Mechanism())
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    if (borderShow) helper.DrawCircle();//绘制圆
                    if (starShow) helper.DrawStar();//绘制星星
                    if (!string.IsNullOrEmpty(companyName)) helper.DrawTitle(companyName);//绘制公司名称
                    if (!string.IsNullOrEmpty(centerText)) helper.DrawHorizontal(centerText);//绘制横向文
                    if (!string.IsNullOrEmpty(bottomText)) helper.DrawChord(bottomText);//绘制下弦文
                                                                                        //helper.Save(Path.Combine(Directory.GetCurrentDirectory(), "公司印章.png"));//保存到本地文件
                    helper.Save(stream, ImageFormat.Png);//保存到流文件
                    byte[] data = new byte[stream.Length];
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.Read(data, 0, Convert.ToInt32(stream.Length));
                    return data;
                }

            }
        }
        /// <summary>
        /// 单位印章转<see cref="byte"/>流,返回文件名
        /// </summary>
        /// <param name="companyName">公司名称</param>
        /// <param name="savePath">保存文件的绝对路径</param>
        /// <param name="centerText">中间文字</param>
        /// <param name="bottomText">下弦文字</param>
        /// <param name="borderShow">是否显示圆形边框</param>
        /// <param name="starShow">是否显示星星</param>
        /// <returns></returns>
        public static string MechanismSaveImgFile(string companyName, string savePath, string centerText = "", string bottomText = "", bool borderShow = true, bool starShow = true)
        {
            using (var helper = SealHelper.Mechanism())
            {
                if (borderShow) helper.DrawCircle();//绘制圆
                if (starShow) helper.DrawStar();//绘制星星
                if (!string.IsNullOrEmpty(companyName)) helper.DrawTitle(companyName,135);//绘制公司名称
                if (!string.IsNullOrEmpty(centerText)) helper.DrawHorizontal(centerText);//绘制横向文
                if (!string.IsNullOrEmpty(bottomText)) helper.DrawChord(bottomText);//绘制下弦文
                string fileName = DateTime.Now.ToString("yyyyMMddHHmmssffffff") + ".png";
                helper.Save(savePath + fileName);//保存到本地文件
                return fileName;
            }
           
        }


        /// <summary>
        /// 个人印章转<see cref="byte"/>流
        /// <paramref name="UserName">用户名</paramref>
        /// <paramref name="borderType">边框类型<see cref="BorderDrawType"/>默认：矩形边框</paramref>
        /// </summary>
        /// <returns></returns>
        public static byte[] PersonalToByte(string UserName, BorderDrawType borderType = BorderDrawType.Rectangle)
        {
            using (var helper = SealHelper.Personal())
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    if (borderType == BorderDrawType.Square)
                    {
                        helper.DrawSquare();//方形印
                        helper.DrawName(UserName);
                    }
                    else if (borderType == BorderDrawType.Rectangle)
                    {
                        helper.DrawNameWithBorder(UserName);//矩形印
                    }
                    else helper.DrawName(UserName);
                    //helper.Save(Path.Combine(Directory.GetCurrentDirectory(), "公司印章.png"));//保存到本地文件
                    helper.Save(stream, ImageFormat.Png);//保存到流文件
                    byte[] data = new byte[stream.Length];
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.Read(data, 0, Convert.ToInt32(stream.Length));
                    return data;
                }

            }
        }
        /// <summary>
        /// 个人印章保存到本地文件夹,返回文件名
        /// </summary>
        /// <paramref name="UserName">用户名</paramref>
        /// <param name="savePath">保存的绝对路径地址</param>
        /// <paramref name="borderType">边框类型<see cref="BorderDrawType"/>默认：矩形边框</paramref>
        /// <returns></returns>
        public static string PersonalSaveImgFile(string UserName, string savePath, BorderDrawType borderType = BorderDrawType.Rectangle)
        {
            using (var helper = SealHelper.Personal())
            {
                if (borderType == BorderDrawType.Square)
                {
                    helper.DrawSquare();//方形印
                    helper.DrawName(UserName);
                }
                else if (borderType == BorderDrawType.Rectangle)
                {
                    helper.DrawNameWithBorder(UserName);//矩形印
                }
                else helper.DrawName(UserName);
                string fileName = DateTime.Now.ToString("yyyyMMddHHmmssffffff") + ".png";
                helper.Save(savePath+ fileName);//保存到本地文件
                return fileName;
            }
           
        }
        /// <summary>
        /// 边框绘制类型
        /// </summary>
        public enum BorderDrawType
        {
            /// <summary>
            /// 无边框
            /// </summary>
            None = 0,
            /// <summary>
            /// 方形
            /// </summary>
            Square = 1,
            /// <summary>
            /// 矩形
            /// </summary>
            Rectangle = 2,
        }
        /// <summary>
        /// 机构印章帮助类
        /// </summary>
        public class MechanismSealHelper : IDisposable
        {
            string star = "★";
            int size = 160;
            Image map;
            Graphics g;
            int defaultWidth;
            float defaultStarSize;
            float defaultTitleSize;
            float defaultHorizontalSize;
            float defaultChordSize;

            public Color Color { get; set; } = Color.Red;
            public string DefaultFontName { get; set; } = "SimSun";

            public MechanismSealHelper()
            {
                map = new Bitmap(size, size);
                g = Graphics.FromImage(map);//实例化Graphics类
                g.SmoothingMode = SmoothingMode.AntiAlias;  //System.Drawing.Drawing2D;           
                g.Clear(Color.Transparent);

                defaultWidth = size / 40;
                defaultStarSize = size / 5;
                defaultTitleSize = (float)Math.Sqrt(size);
                defaultHorizontalSize = (float)Math.Sqrt(size);
                defaultChordSize = size / 20;
            }

            /// <summary>
            /// 绘制圆
            /// </summary>
            public void DrawCircle()
            {
                DrawCircle(defaultWidth);
            }
            /// <summary>
            /// 绘制圆
            /// </summary>
            /// <param name="width">画笔粗细</param>
            public void DrawCircle(int width)
            {
                Rectangle rect = new Rectangle(width, width, size - width * 2, size - width * 2);//设置圆的绘制区域
                Pen pen = new Pen(Color, width);  //设置画笔（颜色和粗细）
                g.DrawEllipse(pen, rect);  //绘制圆
            }
            /// <summary>
            /// 绘制星星
            /// </summary>
            public void DrawStar()
            {
                DrawStar(defaultStarSize, defaultWidth);
            }
            /// <summary>
            /// 绘制星星
            /// </summary>
            /// <param name="emSize"></param>
            public void DrawStar(float emSize)
            {
                DrawStar(emSize, defaultWidth);
            }
            /// <summary>
            /// 绘制星星
            /// </summary>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawStar(float emSize, int width)
            {
                Font starFont = new Font(DefaultFontName, emSize, FontStyle.Bold);//设置星号的字体样式
                var starSize = g.MeasureString(star, starFont);//对指定字符串进行测量
                                                               //要指定的位置绘制星号
                PointF starXy = new PointF(size / 2 - starSize.Width / 2, size / 2 - starSize.Height / 2);
                g.DrawString(star, starFont, new SolidBrush(Color), starXy);
            }

            /// <summary>
            /// 绘制主题
            /// </summary>
            /// <param name="title">主题（公司名称）</param>
            /// <param name="startAngle">开始角度</param>
            public void DrawTitle(string title, float startAngle = 160)
            {
                DrawTitle(title, startAngle, defaultTitleSize);
            }
            /// <summary>
            /// 绘制主题
            /// </summary>
            /// <param name="title">主题（公司名称）</param>
            /// <param name="startAngle">开始角度,必须是左半边，推荐（135-270）</param>
            /// <param name="emSize">字体大小</param>
            public void DrawTitle(string title, float startAngle, float emSize)
            {
                DrawTitle(title, startAngle, emSize, defaultWidth);
            }
            /// <summary>
            /// 绘制主题
            /// </summary>
            /// <param name="title">主题（公司名称）</param>
            /// <param name="startAngle">开始角度,必须是左半边，推荐（135-270）</param>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawTitle(string title, float startAngle, float emSize, int width)
            {
                Font font = new Font(DefaultFontName, emSize, FontStyle.Bold);
                DrawTitle(title, startAngle, font, width);
            }
            /// <summary>
            /// 绘制主题
            /// </summary>
            /// <param name="title">主题（公司名称）</param>
            /// <param name="startAngle">开始角度,必须是左半边，推荐（135-270）</param>
            /// <param name="font">字体</param>
            /// <param name="width">画笔粗细</param>
            public void DrawTitle(string title, float startAngle, Font font, int width)
            {
                if (string.IsNullOrEmpty(title))
                {
                    return;
                }
                if (Math.Cos(startAngle / 180 * Math.PI) > 0)
                {
                    throw new ArgumentException($"初始角度错误：{startAngle}(建议135-270)", nameof(startAngle));
                }

                startAngle = startAngle % 360;//起始角度

                var length = title.Length;
                float changeAngle = (270 - startAngle) * 2 / length;//每个字所占的角度，也就是旋转角度
                var circleWidth = size / 2 - width * 3;//字体圆形的长度
                var fontSize = g.MeasureString(title, font);//测量一下字体
                var per = fontSize.Width / length;//每个字体的长度
                g.SmoothingMode = SmoothingMode.AntiAlias;//消除绘制图形的锯齿
                //起始绘制角度=起始角度+旋转角度/2-字体所占角度的一半
                var angle = startAngle + changeAngle / 2 - (float)(Math.Asin(per / 2 / circleWidth) / Math.PI * 180);//起始绘制角度
                for (var i = 0; i < length; i++)
                {
                    action1(title[i].ToString(), angle, font, width, circleWidth);
                    angle += changeAngle;
                }
            }

            private void action1(string t, float a, Font font, int width, int circleWidth)
            {
                var angleXy = a / 180 * Math.PI;
                var x = size / 2 + Math.Cos(angleXy) * circleWidth;
                var y = size / 2 + Math.Sin(angleXy) * circleWidth;

                DrawChar(t, (float)x, (float)y, a + 90, font, width);
            }
            /// <summary>
            /// 绘制横向文
            /// </summary>
            /// <param name="text">横向文</param>
            public void DrawHorizontal(string text)
            {
                DrawHorizontal(text, defaultHorizontalSize);
            }
            /// <summary>
            /// 绘制横向文
            /// </summary>
            /// <param name="text">横向文</param>
            /// <param name="emSize">字体大小</param>
            public void DrawHorizontal(string text, float emSize)
            {
                Font font = new Font(DefaultFontName, emSize, FontStyle.Bold);//定义字体的字体样式
                DrawHorizontal(text, font);
            }
            /// <summary>
            /// 绘制横向文
            /// </summary>
            /// <param name="text">横向文</param>
            /// <param name="font">字体</param>
            public void DrawHorizontal(string text, Font font)
            {
                int length = text.Length;
                SizeF textSize = g.MeasureString(text, font);//对指定字符串进行测量
                while (textSize.Width > size * 2 / 3)
                {
                    DrawHorizontal(text, new Font(font.Name, font.Size - 1, font.Style));
                    return;
                }
                //要指定的位置绘制中间文字
                PointF point = new PointF(size / 2 - textSize.Width / 2, size * 2 / 3);
                g.SmoothingMode = SmoothingMode.AntiAlias;//消除绘制图形的锯齿
                g.ResetTransform();
                g.DrawString(text, font, new SolidBrush(Color), point);
            }

            /// <summary>
            /// 绘制下弦文
            /// </summary>
            /// <param name="text">下弦文</param>
            /// <param name="startAngle">开始角度,必须是左下半边，推荐（90-180）</param>
            public void DrawChord(string text, float startAngle = 135)
            {
                DrawChord(text, startAngle, defaultChordSize);
            }
            /// <summary>
            /// 绘制下弦文
            /// </summary>
            /// <param name="text">下弦文</param>
            /// <param name="startAngle">开始角度,必须是左下半边，推荐（90-180）</param>
            /// <param name="emSize">字体大小</param>
            public void DrawChord(string text, float startAngle, float emSize)
            {
                DrawChord(text, startAngle, emSize, defaultWidth);
            }
            /// <summary>
            /// 绘制下弦文
            /// </summary>
            /// <param name="text">下弦文</param>
            /// <param name="startAngle">开始角度,必须是左下半边，推荐（90-180）</param>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawChord(string text, float startAngle, float emSize, int width)
            {
                Font font = new Font(DefaultFontName, emSize, FontStyle.Bold);//定义字体的字体样式
                DrawChord(text, startAngle, font, width);
            }
            /// <summary>
            /// 绘制下弦文
            /// </summary>
            /// <param name="text">下弦文</param>
            /// <param name="startAngle">开始角度,必须是左下半边，推荐（90-180）</param>
            /// <param name="font">字体</param>
            /// <param name="width">画笔粗细</param>
            public void DrawChord(string text, float startAngle, Font font, int width)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return;
                }
                if (Math.Cos(startAngle / 180 * Math.PI) > 0)
                {
                    throw new ArgumentException($"初始角度错误：{startAngle}(建议90-135)", nameof(startAngle));
                }
                g.SmoothingMode = SmoothingMode.AntiAlias;//消除绘制图形的锯齿
                startAngle = startAngle % 360;//起始角度

                var length = text.Length;
                float changeAngle = (startAngle - 90) * 2 / length;//每个字所占的角度，也就是旋转角度
                var fontSize = g.MeasureString(text, font);//测量一下字体
                var per = fontSize.Width / length;//每个字体的长度
                var circleWidth = size / 2 - width * 2 - fontSize.Height;//字体圆形的长度

                //起始绘制角度=起始角度-旋转角度/2+字体所占角度的一半
                var angle = startAngle - changeAngle / 2 + (float)(Math.Asin(per / 2 / circleWidth) / Math.PI * 180);//起始绘制角度
                for (var i = 0; i < length; i++)
                {
                    action(text[i].ToString(), angle, font, width, circleWidth);
                    angle -= changeAngle;
                }
            }

            private void action(string t, float a, Font font, int width, float circleWidth)
            {
                var angleXy = a / 180 * Math.PI;
                var x = size / 2 + Math.Cos(angleXy) * circleWidth;
                var y = size / 2 + Math.Sin(angleXy) * circleWidth;

                DrawChar(t, (float)x, (float)y, a - 90, font, width);
            }
            /// <summary>
            /// 绘制单个字符
            /// </summary>
            /// <param name="char">字符</param>
            /// <param name="x">距离画布左边的距离</param>
            /// <param name="y">距离画布上边的距离</param>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawChar(string @char, float x, float y, float emSize)
            {
                DrawChar(@char, x, y, 0, emSize);
            }
            /// <summary>
            /// 绘制单个字符
            /// </summary>
            /// <param name="char">字符</param>
            /// <param name="x">距离画布左边的距离</param>
            /// <param name="y">距离画布上边的距离</param>
            /// <param name="angle">选中角度，0度为右方，顺时针增加</param>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawChar(string @char, float x, float y, float angle, float emSize)
            {
                DrawChar(@char, x, y, angle, emSize, defaultWidth);
            }
            /// <summary>
            /// 绘制单个字符
            /// </summary>
            /// <param name="char">字符</param>
            /// <param name="x">距离画布左边的距离</param>
            /// <param name="y">距离画布上边的距离</param>
            /// <param name="angle">选中角度，0度为右方，顺时针增加</param>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawChar(string @char, float x, float y, float angle, float emSize, int width)
            {
                Font font = new Font(DefaultFontName, emSize, FontStyle.Bold);
                DrawChar(@char, x, y, angle, font, width);
            }
            /// <summary>
            /// 绘制单个字符
            /// </summary>
            /// <param name="char">字符</param>
            /// <param name="x">距离画布左边的距离</param>
            /// <param name="y">距离画布上边的距离</param>
            /// <param name="angle">选中角度，0度为右方，顺时针增加</param>
            /// <param name="fontName">字体</param>
            /// <param name="width">画笔粗细</param>
            public void DrawChar(string @char, float x, float y, float angle, Font font, int width)
            {
                if (string.IsNullOrEmpty(@char) || @char.Length != 1)
                {
                    throw new ArgumentException("only one char is supported", nameof(@char));
                }

                g.ResetTransform();//重置
                g.TranslateTransform(x, y);//调整偏移
                g.RotateTransform(angle);//旋转角度
                g.DrawString(@char, font, new SolidBrush(Color), 0, 0);//绘制，因为使用了偏移，所以这里的坐标是相对偏移的，所以是0
            }

            /// <summary>
            /// 保存图片
            /// </summary>
            /// <param name="fileName"></param>
            public void Save(string fileName)
            {
                map.Save(fileName);
            }
            /// <summary>
            /// 保存图片
            /// </summary>
            /// <param name="stream"></param>
            /// <param name="format"></param>
            public void Save(Stream stream, ImageFormat format)
            {
                map.Save(stream, format);
            }

            public void Dispose()
            {
                try
                {
                    if (map != null)
                    {
                        map.Dispose();
                    }
                    if (g != null)
                    {
                        g.Dispose();
                    }
                }
                catch { }
            }
        }
        /// <summary>
        /// 个人印章帮助类
        /// </summary>
        public class PersonalSealHelper : IDisposable
        {
            int size = 180;
            Image map;
            Graphics g;
            int defaultWidth;
            float defaultSquareSize;
            float defaultNameSize;
            bool isSquare = false;

            public Color Color { get; set; } = Color.Red;
            public string DefaultFontName { get; set; } = "SimSun";

            public PersonalSealHelper()
            {
                map = new Bitmap(size, size);
                g = Graphics.FromImage(map);//实例化Graphics类
                g.SmoothingMode = SmoothingMode.AntiAlias;  //System.Drawing.Drawing2D;           
                g.Clear(Color.Transparent);

                defaultWidth = size / 40;
                defaultSquareSize = size / 4;
                defaultNameSize = size / 4;
            }

            /// <summary>
            /// 绘制方形之印
            /// </summary>
            public void DrawSquare()
            {
                DrawSquare(defaultSquareSize);
            }
            /// <summary>
            /// 绘制方形之印
            /// </summary>
            /// <param name="emSize">字体大小</param>
            public void DrawSquare(float emSize)
            {
                DrawSquare(emSize, defaultWidth);
            }
            /// <summary>
            /// 绘制方形之印
            /// </summary>
            /// <param name="font">字体</param>
            public void DrawSquare(Font font)
            {
                DrawSquare(font, defaultWidth);
            }
            /// <summary>
            /// 绘制方形之印
            /// </summary>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawSquare(float emSize, int width)
            {
                Font font = new Font(DefaultFontName, emSize, FontStyle.Bold);//设置之印的字体样式
                DrawSquare(font, width);
            }
            /// <summary>
            /// 绘制方形之印
            /// </summary>
            /// <param name="font">字体</param>
            /// <param name="width">画笔粗细</param>
            public void DrawSquare(Font font, int width)
            {
                isSquare = true;

                var pen = new Pen(Color, width);//设置画笔的颜色
                Rectangle rect = new Rectangle(width, width, size - width * 2, size - width * 2);//设置绘制区域
                g.DrawRectangle(pen, rect);
                g.SmoothingMode = SmoothingMode.AntiAlias;//消除绘制图形的锯齿
                var textSize = g.MeasureString("之印", font);//对指定字符串进行测量
                var left = (size / 2 - width * 2 - textSize.Width / 2) / 2;
                var perHeght = (size - width * 4 - textSize.Height * 2) / 3;

                PointF point1 = new PointF(left + width * 2, perHeght + width * 2);
                g.DrawString("之", font, pen.Brush, point1);

                PointF point2 = new PointF(left + width * 2, perHeght * 2 + width * 2 + textSize.Height);
                g.DrawString("印", font, pen.Brush, point2);
            }
            /// <summary>
            /// 绘制带边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            public void DrawNameWithBorder(string name)
            {
                DrawNameWithBorder(name, defaultNameSize, defaultWidth);
            }
            /// <summary>
            /// 绘制带边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            /// <param name="font">字体</param>
            public void DrawNameWithBorder(string name, Font font)
            {
                DrawNameWithBorder(name, font, defaultWidth);
            }
            /// <summary>
            /// 绘制带边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            /// <param name="emSize">字体大小</param>
            public void DrawNameWithBorder(string name, float emSize)
            {
                DrawNameWithBorder(name, emSize, defaultWidth);
            }
            /// <summary>
            /// 绘制带边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawNameWithBorder(string name, float emSize, int width)
            {
                Font font = new Font(DefaultFontName, emSize, FontStyle.Bold);//设置字体样式
                DrawNameWithBorder(name, font, width);
            }
            /// <summary>
            /// 绘制带边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            /// <param name="font">字体</param>
            /// <param name="width">画笔粗细</param>
            public void DrawNameWithBorder(string name, Font font, int width)
            {
                var nameSize = g.MeasureString(name, font);//对指定字符串进行测量
                while (nameSize.Width > size - width * 6)
                {
                    DrawNameWithBorder(name, new Font(font.Name, font.Size - 1, font.Style), width);
                    return;
                }
                var left = (int)(size - nameSize.Width - width * 4) / 2;
                var height = (int)(size - nameSize.Height - width * 4) / 2;

                var pen = new Pen(Color, width);//设置画笔的颜色
                Rectangle rect = new Rectangle(width, height, size - width * 2, size - 10 - height * 2);//设置绘制区域
                g.DrawRectangle(pen, rect);

                PointF point = new PointF(width + width * 2, height + width * 2);
                g.DrawString(name, font, pen.Brush, point);
            }
            /// <summary>
            /// 绘制无边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            public void DrawName(string name)
            {
                DrawName(name, defaultNameSize, defaultWidth);
            }
            /// <summary>
            /// 绘制带边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            /// <param name="font">字体</param>
            public void DrawName(string name, Font font)
            {
                DrawName(name, font, defaultWidth);
            }
            /// <summary>
            /// 绘制带边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            /// <param name="emSize">字体大小</param>
            public void DrawName(string name, float emSize)
            {
                DrawName(name, emSize, defaultWidth);
            }
            /// <summary>
            /// 绘制带边框的姓名
            /// </summary>
            /// <param name="name">姓名</param>
            /// <param name="emSize">字体大小</param>
            /// <param name="width">画笔粗细</param>
            public void DrawName(string name, float emSize, int width)
            {
                Font font = new Font(DefaultFontName, emSize, FontStyle.Bold);//设置字体样式
                DrawName(name, font, width);
            }
            /// <summary>
            /// 绘制姓名
            /// </summary>
            /// <param name="name">姓名</param>
            /// <param name="font">字体</param>
            /// <param name="width">画笔粗细</param>
            public void DrawName(string name, Font font, int width)
            {
                var nameSize = g.MeasureString(name, font);//对指定字符串进行测量
                while (nameSize.Width > size - width * 6 || nameSize.Width > (size - width * 6) - 10 && name.Length == 4)//罗文改 || nameSize.Width > (size - width * 6) - 10 && name.Length == 4
                {
                    DrawName(name, new Font(font.Name, font.Size - 1, font.Style), width);
                    return;
                }
                if (isSquare)
                {
                    int length = name.Length;//获取字符串的长度
                    var left = (size / 2 - width - nameSize.Width / length) / 2;
                    //if (left > 20 && name.Length == 4) left = 35;//罗文改
                    var height = (size - width * 4 - nameSize.Height * length) / (length + 1);
                    if (left <= 0 || height <= 0)
                    {
                        return;
                    }

                    for (var i = 0; i < length; i++)
                    {
                        PointF point;
                        if (length == 4) point = new PointF(width + size / 2+20, height * (i + 1) + width * 2 + nameSize.Height * i);
                        else point = new PointF(width + size / 2, height * (i + 1) + width * 2 + nameSize.Height * i);
                        g.SmoothingMode = SmoothingMode.AntiAlias;//消除绘制图形的锯齿
                        g.DrawString(name[i].ToString(), font, new SolidBrush(Color), point);
                    }
                }
                else
                {
                    var left = (size - nameSize.Width) / 2;
                    var height = (size - nameSize.Height) / 2;
                    PointF point = new PointF(width, height);
                    g.SmoothingMode = SmoothingMode.AntiAlias;//消除绘制图形的锯齿
                    g.DrawString(name, font, new SolidBrush(Color), point);
                }
            }

            /// <summary>
            /// 保存图片
            /// </summary>
            /// <param name="fileName"></param>
            public void Save(string fileName)
            {
                map.Save(fileName);
            }
            /// <summary>
            /// 保存图片
            /// </summary>
            /// <param name="stream"></param>
            /// <param name="format"></param>
            public void Save(Stream stream, ImageFormat format)
            {
                map.Save(stream, format);
            }
            public void Dispose()
            {
                try
                {
                    if (map != null)
                    {
                        map.Dispose();
                    }
                    if (g != null)
                    {
                        g.Dispose();
                    }
                }
                catch { }
            }
        }
    }
}
