﻿
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace CommonExpand
{
    /********************************************************************************

    ** 类名称： AsyncLog

    ** 描述：异步记录日志

    ** 引用： System.Web.dll、System.Web.Mvc.dll

    *********************************************************************************/
    /// <summary>
    /// 异步日志
    /// </summary>
    public static class AsyncLog
    {
        /// <summary>
        /// 相对路径
        /// </summary>
        private static readonly string _map = HttpRuntime.AppDomainAppPath + "log/" + DateTime.Now.ToString("yyyy-MM-dd") + " ";

        /// <summary>
        /// 委托方式的异步写入异常
        /// </summary>
        /// <param name="exception">Exception</param>
        /// <param name="context">HttpContext</param>
        private delegate void AsyncLogException(Exception exception, HttpContext context);
        /// <summary>
        /// 委托方式的异步写入关键信息
        /// </summary>
        /// <param name="information">记录的内容</param>
        /// <param name="context">HttpContext</param>
        private delegate void AsyncLogInformation(string information, HttpContext context);
        /// <summary>
        /// 委托方式的异步写入请求
        /// </summary>
        /// <param name="AuthContext">Mvc的<see cref="AuthorizationContext"/>类</param>
        /// <param name="context">HttpContext</param>
        private delegate void AsyncLogMvcRequest(AuthorizationContext AuthContext, HttpContext context);
        /// <summary>
        /// 异步写入异常信息
        /// </summary>
        /// <param name="e">Exception</param>
        private static void BeginLogException(Exception exception, HttpContext context)
        {
            bool flag = exception != null;
            if (flag)
            {
                string text = AsyncLog._map + "error.txt";
                FileInfo fileInfo = new FileInfo(text);
                DirectoryInfo directory = fileInfo.Directory;
                bool flag2 = !directory.Exists;
                if (flag2)
                {
                    directory.Create();
                }
                using (FileStream fileStream = new FileStream(text, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write))
                {
                    StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8);
                    try
                    {
                        streamWriter.BaseStream.Seek(0L, SeekOrigin.End);
                        streamWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        streamWriter.WriteLine("\r\n");
                        streamWriter.WriteLine("\r\n  异常信息：");
                        streamWriter.WriteLine("\r\n\t请求地址：" + context.Request.RawUrl.ToString());
                        streamWriter.WriteLine("\r\n\t错误信息：" + exception.GetRealError());
                        streamWriter.WriteLine("\r\n\t堆栈信息：" + exception.GetStackTrace());
                        streamWriter.WriteLine("\r\n\t错 误 源：" + exception.Source);
                        streamWriter.WriteLine("\r\n");
                        streamWriter.WriteLine("--------------------------------------------------------------------------------------------------------------\n");
                        streamWriter.WriteLine("\r\n");
                        streamWriter.WriteLine("\r\n");
                    }
                    finally
                    {
                        streamWriter.Flush();
                        streamWriter.Dispose();
                    }
                }
            }
        }
        /// <summary>
        /// 异步写入关键信息
        /// </summary>
        /// <param name="information">关键信息</param>
        private static void BeginLogInformation(string information, HttpContext context)
        {
            string text = AsyncLog._map + "key.txt";
            FileInfo fileInfo = new FileInfo(text);
            DirectoryInfo directory = fileInfo.Directory;
            bool flag = !directory.Exists;
            if (flag)
            {
                directory.Create();
            }
            using (FileStream fileStream = new FileStream(text, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write))
            {
                StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8);
                try
                {
                    streamWriter.BaseStream.Seek(0L, SeekOrigin.End);
                    streamWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    streamWriter.WriteLine("\r\n");
                    streamWriter.WriteLine("\r\n");
                    //streamWriter.WriteLine("\r\n\t请求地址：" + context.Request.Url.ToString());
                    streamWriter.WriteLine("\r\n\t请求地址：" + context.Request.Path.ToString());
                    streamWriter.WriteLine("\r\n\t记录信息：" + information);
                    streamWriter.WriteLine("\r\n");
                    streamWriter.WriteLine("--------------------------------------------------------------------------------------------------------------\n");
                    streamWriter.WriteLine("\r\n");
                    streamWriter.WriteLine("\r\n");
                    streamWriter.WriteLine("\r\n");
                }
                finally
                {
                    streamWriter.Flush();
                    streamWriter.Dispose();
                }
            }
        }
        /// <summary>
        /// 异步写入请求信息
        /// </summary>
        /// <param name="AuthContext">AuthorizationContext</param>
        /// <param name="context">HttpContext</param>
        private static void BeginLogMvcRequest(AuthorizationContext AuthContext, HttpContext context)
        {
            string text = AsyncLog._map + "request.txt";
            FileInfo fileInfo = new FileInfo(text);
            DirectoryInfo directory = fileInfo.Directory;
            bool flag = !directory.Exists;
            if (flag)
            {
                directory.Create();
            }
            using (FileStream fileStream = new FileStream(text, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write))
            {
                StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8);
                try
                {
                    streamWriter.BaseStream.Seek(0L, SeekOrigin.End);
                    streamWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    streamWriter.WriteLine("\r\n");
                    streamWriter.WriteLine("\r\n  请求信息：");
                    streamWriter.WriteLine("\r\n\t请求地址：" + context.Request.RawUrl.ToString());
                    streamWriter.WriteLine("\r\n\t请求类型：" + context.Request.RequestType);
                    //streamWriter.WriteLine("\r\n\t请求地址：" + context.Request.Url.ToString());
                    //streamWriter.WriteLine("\r\n\t请求类型：" + context.Request.RequestType);
                    streamWriter.WriteLine("\r\n\t控制器名：" + AuthContext.ActionDescriptor.ControllerDescriptor.ControllerName);
                    streamWriter.WriteLine("\r\n\tAction名：" + AuthContext.ActionDescriptor.ActionName);
                    streamWriter.WriteLine("--------------------------------------------------------------------------------------------------------------\n");
                    streamWriter.WriteLine("\r\n");
                    streamWriter.WriteLine("\r\n");
                    streamWriter.WriteLine("\r\n");
                }
                finally
                {
                    streamWriter.Flush();
                    streamWriter.Dispose();


                    //清除一星期以前的日志文件
                    var Yesterday = DateTime.Today.AddDays(-7);
                    var _FileArr = directory.GetFiles().Where(a => a.CreationTime < Yesterday);
                    foreach (FileInfo file in _FileArr)
                    {
                        try
                        {
                            System.IO.File.Delete(file.FullName);
                        }
                        catch { }
                    }
                }
            }
        }

        /// <summary>
        /// 记录异常信息
        /// </summary>
        /// <param name="exception">Exception</param>
        public static void LogException(Exception exception)
        {  ////异步线程无法访问到主线程的HttpContext，所以要直接将主线程的HttpContext做为参数传给异步
            new AsyncLog.AsyncLogException(AsyncLog.BeginLogException).BeginInvoke(exception, HttpContext.Current, null, null);
        }
        /// <summary>
        /// 记录关键信息
        /// </summary>
        /// <param name="information">记录的内容</param>
        public static void LogInformation(string information)
        {
            new AsyncLog.AsyncLogInformation(AsyncLog.BeginLogInformation).BeginInvoke(information, HttpContext.Current, null, null);
        }
        /// <summary>
        /// 记录请求信息
        /// </summary>
        /// <param name="information">记录的内容</param>
        public static void LogMvcRequest(AuthorizationContext AuthContext)
        {
            new AsyncLog.AsyncLogMvcRequest(AsyncLog.BeginLogMvcRequest).BeginInvoke(AuthContext,HttpContext.Current, null, null);
        }

    }



    //#region 控制器全局异常处理
    ///// <summary>
    ///// 控制器全局异常处理  WebMvc可改
    ///// </summary>
    //public class WebMvc : Controller
    //{
    //    /// <summary>
    //    /// 控制器全局异异常处理
    //    /// </summary>
    //    /// <param name="filterContext"></param>
    //    protected override void OnException(ExceptionContext filterContext)
    //    {
    //        base.OnException(filterContext);
    //        var exception = filterContext.Exception;

    //        //写入异常日志
    //        AsyncLog.LogException(exception);

    //        filterContext.HttpContext.Response.ContentEncoding = Encoding.UTF8;
    //        filterContext.HttpContext.Response.ContentType = "application/json";
    //        //200=正常，400=语法错误，404=无法找到资源
    //        filterContext.HttpContext.Response.StatusCode = 200;
    //        filterContext.HttpContext.Response.Write("{\"code\":12001,\"message\":\"" + exception.GetRealError() + "\"}");
    //        filterContext.HttpContext.Response.End();

    //        //设置异常已处理，不向上一级异常捕获冒泡
    //        filterContext.ExceptionHandled = true;

    //        //清除当前Http中的异常信息，不向View界面抛出
    //        filterContext.HttpContext.ClearError();
    //    }
    //}
    //#endregion
}