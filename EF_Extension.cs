using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Web;
using System.Data.Entity;

namespace CommonExpand
{
    /// <summary>
    /// EF查询延伸类     基于EntityFramework
    /// </summary>
    public static class EF_Extension
    {
        #region EF字段排序

        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public static string SqlConnectionStr { get; set; }

        /// <summary>
        /// 设置指定字段的排序查询数据
        /// </summary>
        /// <param name="query">当前条件语句</param>
        /// <param name="sortFieldName">需要排序的字段名</param>
        /// <param name="order">排序规则(ASC|DESC)，默认：ASC升序</param>
        /// <returns></returns>
        public static IQueryable<T> SetQueryableOrder<T>(this IQueryable<T> query, string sortFieldName, string order = "ASC")
        {
            if (string.IsNullOrEmpty(sortFieldName)) throw new Exception("必须指定排序字段!");
            //根据属性名获取属性
            PropertyInfo sortProperty = typeof(T).GetProperty(sortFieldName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (sortProperty == null) throw new Exception("查询对象中不存在排序字段" + sortFieldName + "！");
            //创建表达式变量参数
            ParameterExpression param = Expression.Parameter(typeof(T), "t");
            Expression body = param;
            if (Nullable.GetUnderlyingType(body.Type) != null)
                body = Expression.Property(body, "Value");
            //创建一个访问属性的表达式
            body = Expression.MakeMemberAccess(body, sortProperty);
            LambdaExpression keySelectorLambda = Expression.Lambda(body, param);
            string queryMethod = order.ToUpper() == "DESC" ? "OrderByDescending" : "OrderBy";
            query = query.Provider.CreateQuery<T>(Expression.Call(typeof(Queryable), queryMethod,
                                                               new Type[] { typeof(T), body.Type },
                                                               query.Expression,
                                                               Expression.Quote(keySelectorLambda)));
            return query;
        }

        /// <summary>
        /// 设置多个指定字段的排序查询数据
        /// </summary>
        /// <param name="query">当前条件语句</param>
        /// <param name="sortFieldString">需要排序的字段,字段间以逗号隔开(例：“FieldA,FieldB”)</param>
        /// <param name="order">排序规则(ASC|DESC)，默认：ASC升序</param>
        /// <returns></returns>
        public static IQueryable<T> SetQueryableOrderArray<T>(this IQueryable<T> query, string sortFieldString, string order = "ASC")
        {
            var sortFieldArray = sortFieldString.Split(',');
            return query.SetQueryableOrderArray(sortFieldArray, order);
        }

        /// <summary>
        /// 设置多个指定字段的排序查询数据
        /// </summary>
        /// <param name="query">当前条件语句</param>
        /// <param name="sortFieldArray">排序字段数组</param>
        /// <param name="order">排序规则(ASC|DESC)，默认：ASC升序</param>
        /// <returns></returns>
        public static IQueryable<T> SetQueryableOrderArray<T>(this IQueryable<T> query, string[] sortFieldArray, string order = "ASC")
        {

            //创建表达式变量参数
            var parameter = Expression.Parameter(typeof(T), "t");

            if (sortFieldArray.Length == 0) throw new Exception("必须指定排序字段!");
            if (sortFieldArray != null && sortFieldArray.Length > 0)
            {
                for (int i = 0; i < sortFieldArray.Length; i++)
                {
                    //根据属性名获取属性
                    var property = typeof(T).GetProperty(sortFieldArray[i], BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                    if (property == null) throw new Exception("查询对象中不存在排序字段" + sortFieldArray[i] + "！");
                    //创建一个访问属性的表达式
                    var propertyAccess = Expression.MakeMemberAccess(parameter, property);
                    var orderByExp = Expression.Lambda(propertyAccess, parameter);


                    string OrderName = order.ToUpper() == "DESC" ? "OrderByDescending" : "OrderBy";


                    MethodCallExpression resultExp = Expression.Call(typeof(Queryable), OrderName, new Type[] { typeof(T), property.PropertyType }, query.Expression, Expression.Quote(orderByExp));
                    query = query.Provider.CreateQuery<T>(resultExp);
                }
            }
            return query;
        }

        #endregion


        #region EF批量插入

        /// <summary>
        /// 批量插入方法
        /// (调用该方法需要注意，字段名称必须和数据库中的字段名称一一对应)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">数据集</param>
        /// <param name="tableName">数据库表名</param>
        /// <param name="connectstring">连接字符串</param>
        public static int BulkInsert<T>(this IList<T> list, string connectstring = "", string tableName = "") where T : class
        {
            if (string.IsNullOrWhiteSpace(connectstring)) connectstring = SqlConnectionStr;
            if (string.IsNullOrWhiteSpace(connectstring)) throw new Exception("连接字符串为空");
            var _tableName = typeof(T).Name;
            SqlConnection conn = new SqlConnection(connectstring);
            using (var bulkCopy = new SqlBulkCopy(conn))
            {
                bulkCopy.BatchSize = list.Count;
                bulkCopy.DestinationTableName = _tableName;

                var table = new DataTable();
                //获取类的属性映射字段，并排除未映射字段
                PropertyInfo[] props = typeof(T).GetType().GetProperties().Where(pi => !Attribute.IsDefined(pi, typeof(System.ComponentModel.DataAnnotations.Schema.NotMappedAttribute))).ToArray();
                //var props = TypeDescriptor.GetProperties(typeof(T), new Attribute[] { })
                //    .Cast<PropertyDescriptor>()
                //    .Where(propertyInfo => propertyInfo.PropertyType.Namespace.Equals("System"))
                //    .ToArray();
                //var props2 = TypeDescriptor.GetProperties(typeof(T), new Attribute[] { new System.ComponentModel.DataObjectAttribute() })//new NotMappedAttribute()
                //       .Cast<PropertyDescriptor>()
                //       .Where(propertyInfo => propertyInfo.PropertyType.Namespace.Equals("System"))
                //       .ToArray();
                //var arr = props2.Select(a => a.Name).ToArray();
                foreach (var propertyInfo in props)
                {
                    bulkCopy.ColumnMappings.Add(propertyInfo.Name, propertyInfo.Name);
                    table.Columns.Add(propertyInfo.Name, Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType);
                }

                var values = new object[props.Length];
                foreach (var item in list)
                {
                    var columns = new List<object>();
                    for (var i = 0; i < values.Length; i++)
                    {
                        //var name = props[i].Name;
                        //if (arr.Contains(name)) continue;
                        columns.Add(props[i].GetValue(item));
                    }
                    var colArr = columns.ToArray();
                    table.Rows.Add(colArr);
                }

                bulkCopy.WriteToServer(table);
                return list.Count;
            }
        }

        /// <summary>
        ///  海量数据插入方法,(调用该方法需要注意，DataTable中的字段名称必须和数据库中的字段名称一一对应)
        /// </summary>
        /// <param name="table">需要插入数据</param>
        /// <param name="connectionString">数据连接字符串</param>
        /// <param name="tableName">目标数据库表的名称</param>
        /// <returns></returns>
        public static int BulkInsert(this DataTable table, string connectionString, string tableName)
        {
            //开始数据保存逻辑
            int totalCount = 0;
            if (string.IsNullOrWhiteSpace(connectionString)) connectionString = SqlConnectionStr;
            if (string.IsNullOrWhiteSpace(connectionString)) throw new Exception("连接字符串为空");
            if (table != null && table.Rows.Count > 0)
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlTransaction tran = conn.BeginTransaction();//开启事务
                    SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.CheckConstraints, tran);//在插入数据的同时检查约束，如果发生错误调用sqlbulkTransaction事务
                    bulkCopy.DestinationTableName = tableName.ToLower();//***代表要插入数据的表名
                    foreach (DataColumn dc in table.Columns)  //传入上述table
                    {
                        bulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);//将table中的列与数据库表这的列一一对应
                    }
                    try
                    {
                        totalCount = table.Rows.Count;
                        bulkCopy.BatchSize = totalCount;//插入的数据条数
                        bulkCopy.BulkCopyTimeout = 120;//超时秒数，120s
                        bulkCopy.WriteToServer(table);
                        tran.Commit();
                    }
                    catch (Exception ex)
                    {
                        tran.Rollback();
                        throw ex;
                    }
                    finally
                    {
                        bulkCopy.Close();
                        conn.Close();
                    }
                }
            }
            return totalCount;
        }
        /// <summary>
        /// 海量数据插入方法,(调用该方法需要注意，DataTable中的字段名称必须和数据库中的字段名称一一对应)
        /// </summary>
        /// <param name="table">需要插入数据</param>
        /// <param name="dbContext">EF上下文</param>
        /// <param name="tableName">目标数据库表的名称</param>
        /// <returns></returns>
        public static int BulkInsert(this DataTable table, DbContext dbContext, string tableName)
        {
            //开始数据保存逻辑
            int totalCount = 0;
            if (table == null || table.Rows.Count == 0) return totalCount;

            using (DbConnection conn = dbContext.Database.Connection)
            {
                conn.Open();
                //DbContextTransaction tran = dbContext.Database.BeginTransaction();//开启事务
                //SqlConnection sqlConnection = (SqlConnection)connection;
                var bulkCopy = new SqlBulkCopy(conn as SqlConnection, SqlBulkCopyOptions.CheckConstraints | SqlBulkCopyOptions.KeepNulls, (dbContext.Database.CurrentTransaction.UnderlyingTransaction) as SqlTransaction);//在插入数据的同时检查约束，如果发生错误调用sqlbulkTransaction事务
                bulkCopy.DestinationTableName = tableName.ToLower();//***代表要插入数据的表名
                foreach (DataColumn dc in table.Columns)  //传入上述table
                {
                    bulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);//将table中的列与数据库表这的列一一对应
                }
                try
                {
                    totalCount = table.Rows.Count;
                    bulkCopy.BatchSize = totalCount;//插入的数据条数
                    bulkCopy.BulkCopyTimeout = 120;//超时秒数，120s
                    bulkCopy.WriteToServer(table);
                }
                catch (Exception ex)
                {
                    dbContext.Database.CurrentTransaction.UnderlyingTransaction.Rollback();
                    throw ex;
                }
                finally
                {
                    bulkCopy.Close();
                    conn.Close();
                }
            }
            return totalCount;
        }
        /// <summary>
        /// 海量数据插入方法,(调用该方法需要注意，DataTable中的字段名称必须和数据库中的字段名称一一对应)
        /// </summary>
        /// <param name="list">需要插入的数据集</param>
        /// <param name="dbContext">EF上下文</param>
        /// <param name="tableName">目标数据库表的名称</param>
        /// <returns></returns>
        public static int BulkInsert<T>(this IList<T> list, DbContext dbContext, string tableName = "")
        {
            //开始数据保存逻辑
            int totalCount = 0;
            if (list == null || list.Count == 0) return totalCount;
            //数据库表明
            string _tableName = typeof(T).Name;
            if (!string.IsNullOrEmpty(tableName)) _tableName = tableName;

            DataTable _table = new DataTable();
            //获取类的属性映射字段，并排除未映射字段
            PropertyInfo[] props = typeof(T).GetType().GetProperties().Where(pi => !Attribute.IsDefined(pi, typeof(System.ComponentModel.DataAnnotations.Schema.NotMappedAttribute))).ToArray();
            //表头
            foreach (var propertyInfo in props)
            {
                _table.Columns.Add(propertyInfo.Name, Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType);
            }
            //表数据
            var values = new object[props.Length];
            foreach (var item in list)
            {
                var columns = new List<object>();
                for (var i = 0; i < values.Length; i++)
                {
                    columns.Add(props[i].GetValue(item));
                }
                var colArr = columns.ToArray();
                _table.Rows.Add(colArr);
            }
            return _table.BulkInsert(dbContext, _tableName);
        }

        #endregion


    }
}