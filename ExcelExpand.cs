using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace ExpandComponents
{
    /********************************************************************************

    ** 类名称： ExcelExpand

    ** 描述：导入导出Excel拓展

    ** 引用： NPOI.dll   NPOI.OOML.dll  NPOI.OpenXml4Net.dll  NPOI.OpenXmlFormats.dll  ICSharpCode.SharpZipLib.dll(版本：1.0.0.999)

    ** 作者： LW

    *********************************************************************************/

    /// <summary>
    /// 导入导出Excel拓展
    /// </summary>
    public static class ExcelExpand
    {
        #region 导入读取文件拓展

        #region Microsoft.ACE.OLEDB.12.0 导入

        /// <summary>
        /// 将Excel数据表格转换为DateTable,返回excel第一个有数据的表
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="del">是否删除原文件</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string path, bool del = false)
        {
            bool flag = !File.Exists(path);
            if (flag)
            {
                throw new Exception("未找到 path 中指定的文件。");
            }
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
            DataSet dataSet = new DataSet();
            DataTable result;
            try
            {
                oleDbConnection.Open();
                DataTable oleDbSchemaTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                DataTable dataTable = new DataTable();
                for (int i = 0; i < oleDbSchemaTable.Rows.Count; i++)
                {
                    string value = oleDbSchemaTable.Rows[i]["TABLE_NAME"].ToString();
                    bool flag2 = string.IsNullOrEmpty(value);
                    if (!flag2)
                    {
                        string selectCmddText = "select * from [" + value + "]";
                        OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(selectCmddText, oleDbConnection);
                        oleDbDataAdapter.Fill(dataSet, "excelData");
                        var count = dataSet.Tables[0].Rows.Count;
                        if (count > 0) { dataTable = dataSet.Tables[0]; break; };
                    }
                }
                for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
                {
                    int num = 0;
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        string value = dataTable.Rows[i][j].ToString();
                        bool flag2 = string.IsNullOrEmpty(value);
                        if (flag2)
                        {
                            num++;
                        }
                    }
                    bool flag3 = num == dataTable.Columns.Count;
                    if (flag3)
                    {
                        dataTable.Rows.RemoveAt(i);
                    }
                }
                result = dataTable;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                bool flag4 = oleDbConnection.State > ConnectionState.Closed;
                if (flag4)
                {
                    oleDbConnection.Close();
                }
                oleDbConnection.Dispose();
                if (del)
                {
                    if (File.Exists(path)) File.Delete(path);
                }
            }
            return result;
        }
        /// <summary>
        /// 将Excel数据表格转换为DateTable，指定表名
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="TableName">Excel表名</param>
        /// <param name="del">是否删除原文件</param>
        /// <returns></returns>
        public static DataTable ExcelToDataTable(string path, string TableName, bool del = false)
        {
            bool flag = !File.Exists(path);
            if (flag)
            {
                throw new Exception("未找到 path 中指定的文件。");
            }
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
            DataSet dataSet = new DataSet();
            DataTable result;
            try
            {
                oleDbConnection.Open();
                DataTable oleDbSchemaTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                DataTable dataTable = new DataTable();
                string selectCmddText = "select * from [" + TableName + "]";
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(selectCmddText, oleDbConnection);
                oleDbDataAdapter.Fill(dataSet, "excelData");
                var count = dataSet.Tables[0].Rows.Count;
                if (count > 0) { dataTable = dataSet.Tables[0]; }
                for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
                {
                    int num = 0;
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        string value = dataTable.Rows[i][j].ToString();
                        bool flag2 = string.IsNullOrEmpty(value);
                        if (flag2)
                        {
                            num++;
                        }
                    }
                    bool flag3 = num == dataTable.Columns.Count;
                    if (flag3)
                    {
                        dataTable.Rows.RemoveAt(i);
                    }
                }
                result = dataTable;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                bool flag4 = oleDbConnection.State > ConnectionState.Closed;
                if (flag4)
                {
                    oleDbConnection.Close();
                }
                oleDbConnection.Dispose();
                if (del)
                {
                    if (File.Exists(path)) File.Delete(path);
                }
            }
            return result;
        }

        #endregion


        #region NPOI 导入

        #region 将 Excel 文件读取到 DataTable
        /// <summary>
        /// 将 Excel 文件读取到 <see cref="DataTable"/>
        /// </summary>
        /// <param name="filePath">文件完整路径名</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="firstRowIsColumnName">首行是否为 <see cref="DataColumn.ColumnName"/></param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadExcelToDataTable(string filePath, string sheetName = null)
        {
            if (string.IsNullOrEmpty(filePath)) return null;
            if (!File.Exists(filePath)) return null;

            //根据指定路径读取文件
            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (fileStream == null || fileStream.Length <= 0) return null;
            //Excel行号
            var RowIndex = 1;
            try
            {
                var data = ReadStreamData(fileStream, sheetName, 0, ref RowIndex);
                return data;
            }
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                fileStream.Close(); // 关闭流
                fileStream.Dispose();
            }
        }
        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="DataTable"/>,指定表头索引
        /// </summary>
        /// <param name="stream">当前 <see cref="Stream"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="HeaderIndex">指定表头行索引,默认第一行索引为:0</param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadExcelToDataTable_Specifies_Header(string filePath, string sheetName = null, int HeaderIndex = 0)
        {
            if (string.IsNullOrEmpty(filePath)) return null;
            if (!File.Exists(filePath)) return null;

            //根据指定路径读取文件
            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (fileStream == null || fileStream.Length <= 0) return null;
            //Excel行号
            var RowIndex = 1;
            try
            {
                var data = ReadStreamData(fileStream, sheetName, HeaderIndex, ref RowIndex);
                return data;
            }
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                fileStream.Close(); // 关闭流
                fileStream.Dispose();
            }
        }
        #endregion

        #region 将 Stream 对象读取到 DataTable
        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="DataTable"/>
        /// </summary>
        /// <param name="stream">当前 <see cref="Stream"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="firstRowIsColumnName">首行是否为 <see cref="DataColumn.ColumnName"/></param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadStreamToDataTable(Stream stream, string sheetName = null)
        {
            var table = new DataTable();
            if (stream == null || !stream.CanRead || stream.Length <= 0) return table;
            //Excel行号
            var RowIndex = 1;
            try
            {
                var data = ReadStreamData(stream, sheetName,0,ref RowIndex);
                return data;
            }
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                stream.Close(); // 关闭流
                stream.Dispose();
            }
        }

        private static DataTable ReadStreamData(Stream stream, string sheetName,int HeaderIndex, ref int RowIndex)
        {
            //定义要返回的datatable对象
            var data = new DataTable();
            //数据开始行(排除标题行)
            int startRow = 0;
            ISheet sheet;
            //根据文件流创建excel数据结构,NPOI的工厂类WorkbookFactory会自动识别excel版本，创建出不同的excel数据结构
            IWorkbook workbook = WorkbookFactory.Create(stream);

            //如果有指定工作表名称
            if (!string.IsNullOrEmpty(sheetName))
            {
                sheet = workbook.GetSheet(sheetName);
                //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                if (sheet == null)
                {
                    sheet = workbook.GetSheetAt(0);
                }
            }
            else
            {
                //如果没有指定的sheetName，则尝试获取第一个sheet
                sheet = workbook.GetSheetAt(0);
            }
            if (sheet != null)
            {
                IRow firstRow = sheet.GetRow(HeaderIndex);
                if (firstRow == null || firstRow.FirstCellNum < 0) throw new Exception("未获取到表头数据");
                RowIndex = firstRow.RowNum + 1;
                //一行最后一个cell的编号 即总的列数
                int cellCount = firstRow.LastCellNum;

                //如果第一行是标题列名
                if (HeaderIndex >= 0)
                {
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        ICell cell = firstRow.GetCell(i);
                        if (cell != null)
                        {
                            string cellValue = cell.StringCellValue;
                            if (cellValue != null)
                            {
                                cellValue = cellValue.Trim().Replace(" ", "");
                                if (data.Columns.Contains(cellValue)) throw new Exception($"表头存在多个列名为“{cellValue}”的列！");
                                DataColumn column = new DataColumn(cellValue);
                                data.Columns.Add(column);
                            }
                            else
                            {
                                DataColumn column = new DataColumn("Column" + (i + 1));
                                data.Columns.Add(column);
                            }
                        }
                        else
                        {
                            DataColumn column = new DataColumn("Column" + (i + 1));
                            data.Columns.Add(column);
                        }
                    }
                    if (cellCount > 0)
                    {
                        DataColumn column = new DataColumn("RowNum");
                        data.Columns.Add(column);
                    }
                    startRow = HeaderIndex + 1;
                }

                //最后一列的标号
                int rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null || row.FirstCellNum < 0) continue; //没有数据的行默认是null　
                    RowIndex = row.RowNum + 1;
                    var blankCount = 0;//空白单元格数
                    DataRow dataRow = data.NewRow();
                    for (int j = row.FirstCellNum; j < row.FirstCellNum + cellCount; ++j)
                    {
                        //同理，没有数据的单元格都默认是null
                        ICell cell = row.GetCell(j);
                        //判断单元格是否为空白
                        if (cell == null || cell.CellType == CellType.Blank) { blankCount++; continue; }
                        if (cell != null)
                        {
                            if (cell.CellType == CellType.Numeric)
                            {
                                short format = cell.CellStyle.DataFormat;//是否是带格式的日期类型
                                if (format == 0xe || format == 0x16) { dataRow[j] = cell.DateCellValue.ToString(); }
                                else { dataRow[j] =cell.NumericCellValue.ToString(); }
                                //dataRow[j] = row.GetCell(j).ToString().Trim();
                                //判断是否日期类型
                                //if (DateUtil.IsCellDateFormatted(cell))
                                //{
                                //    dataRow[j] = row.GetCell(j).DateCellValue;
                                //}
                                //else
                                //{
                                //    dataRow[j] = row.GetCell(j).ToString().Trim();
                                //}
                            }
                            else if (cell.CellType == CellType.Formula)
                            {
                                CellType CellType = row.GetCell(j).CachedFormulaResultType;
                                //判断是否公式计算值
                                if (CellType == CellType.String)
                                    dataRow[j] = row.GetCell(j).StringCellValue.ToString().Trim();
                                else if (CellType == CellType.Numeric)
                                    dataRow[j] = row.GetCell(j).NumericCellValue.ToString().Trim();
                                else if (CellType == CellType.Blank) dataRow[j] = "";
                                else dataRow[j] = row.GetCell(j).ToString().Trim();
                            }
                            else
                            {
                                dataRow[j] = row.GetCell(j).ToString().Trim();
                            }
                        }
                    }
                    if (blankCount == cellCount) continue;
                    if (cellCount > 0)
                    {
                        var ColumnIndex = data.Columns.Count - 1;
                        dataRow[ColumnIndex] = RowIndex.ToString();
                    }
                    data.Rows.Add(dataRow);
                }
            }

            return data;
        }

        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="DataTable"/>,指定表头索引,含 RowNum 行号
        /// </summary>
        /// <param name="stream">当前 <see cref="Stream"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="HeaderIndex">指定表头行索引,默认第一行索引为:0</param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadStreamToDataTable_Specifies_Header(Stream stream, string sheetName = null, int HeaderIndex = 0)
        {
            if (stream == null || stream.Length <= 0) return null;
            //Excel行号
            var RowIndex = 1;
            try
            {
                var data = ReadStreamData(stream, sheetName, HeaderIndex, ref RowIndex);
                return data;
            }
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                stream.Close(); // 关闭流
                stream.Dispose();
            }
        }
        //判断某行某列有问题
        private static int CheckRowError(HSSFCell cell)
        {
            //判断各个单元格是否为空
            if (cell == null || cell.Equals(""))
            {
                return -1;
            }
            return 0;
        }
        #endregion

        /// <summary>
        /// 是否为Excel文件
        /// </summary>
        /// <returns></returns>
        public static bool FileIsExcel(string path)
        {
            var _fileInfo = new FileInfo(path);
            if (_fileInfo == null) return false;
            var ext = _fileInfo.Extension.ToLower();
            if (ext == ".xls" || ext == ".xlsx") return true;
            return false;
        }

        #endregion

        #endregion

        #region 导出Excel拓展

        #region NPOI 导出 单元格拓展
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列,第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, string Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (!string.IsNullOrEmpty(Value)) cell.SetCellValue(Value);
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, int? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, decimal? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, decimal? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, double? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, double? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, float? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, float? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, ref int Index, bool? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">时间格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, DateTime? Value, string Format = "yyyy-MM-dd")
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                Index++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, ref int Index, string Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, int? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, decimal? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, decimal? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, double? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, double? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, float? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, float? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, bool? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, ref int Index, DateTime? Value, ICellStyle Style, string Format = "yyyy-MM-dd")
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }



        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列</param>
        /// <param name="Value">值</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, string Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (!string.IsNullOrEmpty(Value)) cell.SetCellValue(Value);
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, int? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, decimal? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, decimal? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, double? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, double? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, float? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, float? Value, string Format)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public static ICell CreateTd(this IRow Row, int Index, bool? Value)
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.ToString());
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Format">时间格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, DateTime? Value, string Format = "yyyy-MM-dd")
        {
            try
            {
                var cell = Row.CreateCell(Index);
                if (Value != null) cell.SetCellValue(Value.Value.ToString(Format));
                
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Row">当前<see cref="IRow"/>对象</param>
        /// <param name="Index">列索引，第X列，第一列必须为零，自动返回下一个列索引</param>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public static ICell CreateTd(this IRow Row, int Index, string Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, int? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, decimal? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, decimal? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, double? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, double? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, float? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, float? Value, string Format, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, bool? Value, ICellStyle Style)
        {
            var cell = Row.CreateTd(ref Index, Value);
            cell.CellStyle = Style;
            return cell;
        }
        public static ICell CreateTd(this IRow Row, int Index, DateTime? Value, ICellStyle Style, string Format = "yyyy-MM-dd")
        {
            var cell = Row.CreateTd(ref Index, Value, Format);
            cell.CellStyle = Style;
            return cell;
        }

        
        #endregion


        /*

        #region 制作表格
               //表格制作
                var ExcelName="tableExport.xls";//excel文件名(含后缀)

                var book = new HSSFWorkbook();
                var sheet = book.CreateSheet("Sheet1");
                //表头样式
                ICellStyle HeadStyle = book.CreateCellStyle();
                HSSFFont font = (HSSFFont)book.CreateFont();
                font.FontName = "黑体";//字体
                font.Boldweight = 700;//加粗
                font.Color = HSSFColor.Black.Index;//颜色
                CellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中 
                CellStyle.Alignment = HorizontalAlignment.CenterSelection;//水平居中
                HeadStyle.SetFont(font);

                 //单元格样式
                ICellStyle CellStyle = book.CreateCellStyle();
                CellStyle.VerticalAlignment = VerticalAlignment.Center;//垂直居中 
                CellStyle.Alignment = HorizontalAlignment.CenterSelection;//水平居中
                CellStyle.SetFont(book.CreateFont());

                //合并单元格
                //sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 12));

                //创建表头数组
                 var head = new string[] {
                    "利润中心", "县支公司代码", "对同一供应商的付款上限值", "年份", "方案号","类别",
                    "修改日期","修改人工号"
                };
                var rowIndex = 0; //行索引
                var hrow = sheet.CreateRow(rowIndex);//创建表头
                for (int i = 0; i < head.Length; i++)
                {
                    var columnIndex = i;
                    hrow.CreateTd(ref columnIndex, head[i],HeadStyle);
                }
                //循环数据列表
                foreach (var item in list)
                {
                    rowIndex++;
                    var trow = sheet.CreateRow(rowIndex); 
                    var columnIndex = 0;//列索引
                    trow.CreateTd(ref columnIndex, item.branch_no,CellStyle);//创建单元格
                }

                //自动列宽
                for (int i = 0; i < head.Length; i++)
                    sheet.AutoSizeColumn(i, true);
               //输出文件流
                using (var ms = new MemoryStream())
                {
                    book.Write(ms);
                    var value = ms.ToArray();

                    book.Close();
                    return File(value, "application/vnd.ms - excel",ExcelName);
                }
                 #endregion

        */

        #endregion
    }

    /********************************************************************************

** 类名称： ExcelExpand

** 描述：导入导出Excel拓展

** 引用：(Nuget可找到) NPOI.dll   NPOI.OOML.dll  NPOI.OpenXml4Net.dll  NPOI.OpenXmlFormats.dll  ICSharpCode.SharpZipLib.dll

** 作者： LW

*********************************************************************************/

    /// <summary>
    /// Excel导入拓展
    /// </summary>
    public class ExcelImport
    {
        #region 获取单个Sheet表格，将 Stream 对象读取到 DataTable
        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="DataTable"/>，获取单个Sheet表格
        /// </summary>
        /// <param name="stream">要读取的 <see cref="Stream"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <param name="dispose">是否释放 <see cref="Stream"/> 资源</param>
        /// <returns>
        /// 如果 stream 参数为 null，则返回 null；
        /// 如果 stream 参数的 <see cref="Stream.CanRead"/> 属性为 false，则返回 null；
        /// 如果 stream 参数的 <see cref="Stream.Length"/> 属性为 小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="Stream"/> 读取后的 <see cref="DataTable"/> 对象。
        /// </returns>
        public DataTable ReadStreamToDataTable(Stream stream, string sheetName = null, int headRowIndex = 0, bool addEmptyRow = false, bool dispose = true)
        {
            var table = new DataTable();
            if (stream == null || !stream.CanRead || stream.Length <= 0) return table;
            var workbook = WorkbookFactory.Create(stream);
            var sheet = workbook.GetSheetAt(0);
            if (!string.IsNullOrEmpty(sheetName)) sheet = workbook.GetSheet(sheetName);
            if (sheet == null) return table;

            table = ReadSheetToDataTable(sheet, headRowIndex, addEmptyRow);
            if (dispose)
            {
                //stream.Flush();
                //stream.Close();
            }
            return table;
        }
        #endregion

        #region 获取多个Sheet表格
        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="ICollection{DataTable}"/>,返回Excel所有的Sheet表格,用于获取多个表格
        /// </summary>
        /// <param name="stream">要读取的 <see cref="Stream"/> 对象</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="ignoreSheetName">忽略的Sheet表名或索引,忽略多个表用英文逗号隔开。例：Sheet1,Sheet2或0,1</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <param name="dispose">是否释放 <see cref="Stream"/> 资源</param>
        /// <returns>
        /// 如果 stream 参数为 null，则返回 null；
        /// 如果 stream 参数的 <see cref="Stream.CanRead"/> 属性为 false，则返回 null；
        /// 如果 stream 参数的 <see cref="Stream.Length"/> 属性小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="Stream"/> 读取后的 <see cref="ICollection{DataTable}"/> 对象，
        /// 其中一个 <see cref="DataTable"/> 对应一个 Sheet 工作簿。
        /// </returns>
        public ICollection<DataTable> ReadStreamToTables(Stream stream, int headRowIndex = 0, string ignoreSheetName = "", bool addEmptyRow = false, bool dispose = true)
        {
            if (stream == null || !stream.CanRead || stream.Length <= 0) return null;
            var tables = new HashSet<DataTable>();

            var workbook = WorkbookFactory.Create(stream);
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                if (sheet == null) continue;
                if (!string.IsNullOrEmpty(ignoreSheetName))
                {
                    var ignoreSheetArr = ignoreSheetName.Split(',');
                    if (ignoreSheetArr.Contains(sheet.SheetName)) continue;
                    if (ignoreSheetArr.Contains(i.ToString())) continue;
                }
                var dataTable = ReadSheetToDataTable(sheet, headRowIndex, addEmptyRow);
                tables.Add(dataTable);
            }

            if (dispose)
            {
                stream.Flush();
                stream.Close();
            }
            return tables;
        }

        /// <summary>
        /// 将 <see cref="Stream"/> 用异步方式读取到 <see cref="ICollection{DataTable}"/>,返回Excel所有的Sheet表格,用于获取多个表格,获取的<see cref="DataTable"/>无序,Sheet越小越容易读取出来
        /// </summary>
        /// <param name="httpPostedFile">要读取的 <see cref="HttpPostedFileBase"/> 对象</param>
        /// <param name="firstRowIsColumnName">首行是否为 <see cref="DataColumn.ColumnName"/></param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 httpPostedFile 参数为 null，则返回 null；
        /// 如果 httpPostedFile 参数的 <see cref="HttpPostedFileBase.ContentLength"/> 属性小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="HttpPostedFileBase"/>读取后的 <see cref="ICollection{DataTable}"/> 对象，
        /// 其中一个 <see cref="DataTable"/> 对应一个 Sheet 工作簿。
        /// </returns>
        public async Task<ICollection<DataTable>> ReadStreamToTablesAsync_plus(Stream stream, int headRowIndex = 0, bool addEmptyRow = false, bool dispose = true)
        {

            if (stream == null || !stream.CanRead || stream.Length <= 0) return null;
            var tables = new HashSet<DataTable>();
            var workbook = WorkbookFactory.Create(stream);
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                if (sheet == null) continue;
                var dataTable = await Task.Run(() => ReadSheetToDataTable(sheet, headRowIndex, addEmptyRow));
                tables.Add(dataTable);
            }

            if (dispose)
            {
                stream.Flush();
                stream.Close();
            }
            return tables;

        }

        /// <summary>
        /// 将 <see cref="Stream"/> 用异步方式读取到 <see cref="ICollection{DataTable}"/>,返回Excel所有的Sheet表格,用于获取多个表格,按顺序依次获取
        /// </summary>
        /// <param name="stream">要读取的 <see cref="Stream"/> 对象</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="ignoreSheetName">忽略的Sheet表名或索引,忽略多个表用英文逗号隔开。例：Sheet1,Sheet2或索引 0,1</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 httpPostedFile 参数为 null，则返回 null；
        /// 如果 httpPostedFile 参数的 <see cref="HttpPostedFileBase.ContentLength"/> 属性小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="HttpPostedFileBase"/>读取后的 <see cref="ICollection{DataTable}"/> 对象，
        /// 其中一个 <see cref="DataTable"/> 对应一个 Sheet 工作簿。
        /// </returns>
        public async Task<ICollection<DataTable>> ReadStreamToTablesAsync(Stream stream, int headRowIndex = 0, string ignoreSheetName = "", bool addEmptyRow = false,
            bool dispose = true) => await Task.Run(() => ReadStreamToTables(stream, headRowIndex, ignoreSheetName, addEmptyRow, dispose));
        #endregion


        #region 将指定路径的 Excel 文件读取到 DataTable
        /// <summary>
        /// 将指定路径的 Excel 文件读取到 <see cref="DataTable"/>
        /// </summary>
        /// <param name="filePath">指定文件完整路径名,文件绝对路径</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 filePath 参数为 null 或者空字符串("")，则返回 null；
        /// 如果 filePath 参数值的磁盘中不存在 Excel 文件，则返回 null；
        /// 否则返回从指定 Excel 文件读取后的 <see cref="DataTable"/> 对象。
        /// </returns>
        public DataTable ReadExcelToDataTable(string filePath, string sheetName = null, int headRowIndex = 0, bool addEmptyRow = false)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return new DataTable();
            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var dt = ReadStreamToDataTable(fileStream, sheetName, headRowIndex, addEmptyRow);
            return dt;
        }

        /// <summary>
        /// 将指定路径的 Excel 文件用异步方式读取到 <see cref="DataTable"/>
        /// </summary>
        /// <param name="filePath">指定文件完整路径名</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="headRowIndex">表头行索引，默认：0,第一行</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 filePath 参数为 null 或者空字符串("")，则返回 null；
        /// 如果 filePath 参数值的磁盘中不存在 Excel 文件，则返回 null；
        /// 否则返回从指定 Excel 文件读取后的 <see cref="DataTable"/> 对象。
        /// </returns>
        public async Task<DataTable> ReadExcelToDataTableAsync(string filePath, string sheetName = null, int headRowIndex = 0, bool addEmptyRow = false) =>
            await Task.Run(() => ReadExcelToDataTable(filePath, sheetName, headRowIndex, addEmptyRow));
        #endregion

        #region 将指定路径的Excel文件读取到 ICollection<DataTable>
        /// <summary>
        /// 将指定路径的 Excel 文件读取到 <see cref="ICollection{DataTable}"/>
        /// </summary>
        /// <param name="filePath">指定文件完整路径名,文件绝对路径</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="ignoreSheetName">忽略的Sheet表名或索引,忽略多个表用英文逗号隔开。例：Sheet1,Sheet2或索引 0,1</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 filePath 参数为 null 或者空字符串("")，则返回 null；
        /// 如果 filePath 参数值的磁盘中不存在 Excel 文件，则返回 null；
        /// 否则返回从指定 Excel 文件读取后的 <see cref="ICollection{DataTable}"/> 对象，
        /// 其中一个 <see cref="DataTable"/> 对应一个 Sheet 工作簿。
        /// </returns>
        public ICollection<DataTable> ReadExcelToTables(string filePath, int headRowIndex = 0, string ignoreSheetName = "", bool addEmptyRow = false)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return null;

            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var tables = ReadStreamToTables(fileStream, headRowIndex, ignoreSheetName, addEmptyRow);
            return tables;
        }
        /// <summary>
        /// 将指定路径的 Excel 文件用异步方式读取到 <see cref="ICollection{DataTable}"/>
        /// </summary>
        /// <param name="filePath">指定文件完整路径名</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="ignoreSheetName">忽略的Sheet表名或索引,忽略多个表用英文逗号隔开。例：Sheet1,Sheet2或索引 0,1</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 filePath 参数为 null 或者空字符串("")，则返回 null；
        /// 如果 filePath 参数值的磁盘中不存在 Excel 文件，则返回 null；
        /// 否则返回从指定 Excel 文件读取后的 <see cref="ICollection{DataTable}"/> 对象，
        /// 其中一个 <see cref="DataTable"/> 对应一个 Sheet 工作簿。
        /// </returns>
        public async Task<ICollection<DataTable>> ReadExcelToTablesAsync(string filePath, int headRowIndex = 0, string ignoreSheetName = "", bool addEmptyRow = false) =>
            await Task.Run(() => ReadExcelToTables(filePath, headRowIndex, ignoreSheetName, addEmptyRow));
        #endregion


        #region 将 HttpPostedFileBase读取到DataTable

        /// <summary>
        /// 将 <see cref="HttpPostedFileBase"/> 读取到 <see cref="DataTable"/>
        /// </summary>
        /// <param name="httpPostedFile">要读取的 <see cref="HttpPostedFileBase"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 httpPostedFile 参数为 null，则返回 null；
        /// 如果 httpPostedFile 参数的 <see cref="HttpPostedFileBase.ContentLength"/> 属性为 小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="HttpPostedFileBase"/>读取后的 <see cref="DataTable"/> 对象。
        /// </returns>
        public DataTable ReadHttpPostedFileToDataTable(HttpPostedFileBase httpPostedFile, string sheetName = null, int headRowIndex = 0, bool addEmptyRow = false)
        {
            if (httpPostedFile == null || httpPostedFile.ContentLength <= 0) return new DataTable();
            var dataTable = ReadStreamToDataTable(httpPostedFile.InputStream, sheetName, headRowIndex, addEmptyRow, dispose: false);
            return dataTable;
        }
        /// <summary>
        /// 将 <see cref="HttpPostedFileBase"/> 用异步方式读取到 <see cref="DataTable"/>
        /// </summary>
        /// <param name="httpPostedFile">要读取的 <see cref="HttpPostedFileBase"/> 对象</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 httpPostedFile 参数为 null，则返回 null；
        /// 如果 httpPostedFile 参数的 <see cref="HttpPostedFileBase.ContentLength"/> 属性为 小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="HttpPostedFileBase"/>读取后的 <see cref="DataTable"/> 对象。
        /// </returns>
        public async Task<DataTable> ReadHttpPostedFileToDataTableAsync(HttpPostedFileBase httpPostedFile, string sheetName = null, int headRowIndex = 0,
            bool addEmptyRow = false) => await Task.Run(() => ReadHttpPostedFileToDataTable(httpPostedFile, sheetName, headRowIndex, addEmptyRow));

        #endregion

        #region 将 HttpPostedFileBase 读取到 ICollection<DataTable>

        /// <summary>
        /// 将 <see cref="HttpPostedFileBase"/> 读取到 <see cref="ICollection{DataTable}"/>
        /// </summary>
        /// <param name="httpPostedFile">要读取的 <see cref="HttpPostedFileBase"/> 对象</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="ignoreSheetName">忽略的Sheet表名或索引,忽略多个表用英文逗号隔开。例：Sheet1,Sheet2或索引 0,1</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 httpPostedFile 参数为 null，则返回 null；
        /// 如果 httpPostedFile 参数的 <see cref="HttpPostedFileBase.ContentLength"/> 属性小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="HttpPostedFileBase"/>读取后的 <see cref="ICollection{DataTable}"/> 对象，
        /// 其中一个 <see cref="DataTable"/> 对应一个 Sheet 工作簿。
        /// </returns>
        public ICollection<DataTable> ReadHttpPostedFileToTables(HttpPostedFileBase httpPostedFile, int headRowIndex = 0, string ignoreSheetName = "", bool addEmptyRow = false)
        {
            if (httpPostedFile == null || httpPostedFile.ContentLength <= 0) return null;

            var tables = ReadStreamToTables(httpPostedFile.InputStream, headRowIndex, ignoreSheetName, addEmptyRow, dispose: false);
            return tables;
        }

        /// <summary>
        /// 将 <see cref="HttpPostedFileBase"/> 用异步方式读取到 <see cref="ICollection{DataTable}"/>
        /// </summary>
        /// <param name="httpPostedFile">要读取的 <see cref="HttpPostedFileBase"/> 对象</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="ignoreSheetName">忽略的Sheet表名或索引,忽略多个表用英文逗号隔开。例：Sheet1,Sheet2或索引 0,1</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 httpPostedFile 参数为 null，则返回 null；
        /// 如果 httpPostedFile 参数的 <see cref="HttpPostedFileBase.ContentLength"/> 属性小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="HttpPostedFileBase"/>读取后的 <see cref="ICollection{DataTable}"/> 对象，
        /// 其中一个 <see cref="DataTable"/> 对应一个 Sheet 工作簿。
        /// </returns>
        public async Task<ICollection<DataTable>> ReadHttpPostedFileToTablesAsync(HttpPostedFileBase httpPostedFile, int headRowIndex = 0, string ignoreSheetName = "",
            bool addEmptyRow = false) => await Task.Run(() => ReadHttpPostedFileToTables(httpPostedFile, headRowIndex, ignoreSheetName, addEmptyRow));
        #endregion

        #region 多Excel文件上传
        /// <summary>
        /// 将 <see cref="HttpFileCollectionBase"/> 对象读取到 <see cref="ICollection{Collection}"/> 集合
        /// </summary>
        /// <param name="httpFileCollection">要读取的 <see cref="HttpFileCollectionBase"/></param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns>
        /// 如果 httpFileCollection 参数为 null，则返回 null；
        /// 如果 httpFileCollection 参数的 <see cref="HttpFileCollectionBase.Count"/> 属性小于或者等于 0，则返回 null；
        /// 否则返回从 <see cref="HttpFileCollectionBase"/> 读取后的 <see cref="ICollection{Collection}"/> 集合。
        /// 结构说明：
        /// <see cref="HttpPostedFileBase"/> 对应 <see cref="Collection{DataTable}"/>;
        /// <see cref="DataTable"/> 对应 Sheet 工作簿。
        /// </returns>
        public ICollection<List<DataTable>> ReadHttpFileCollectionToTableCollection(HttpFileCollectionBase httpFileCollection, int headRowIndex = 0, bool addEmptyRow = false)
        {
            if (httpFileCollection == null || httpFileCollection.Count <= 0) return null;

            var collection = new HashSet<List<DataTable>>();
            for (int i = 0; i < httpFileCollection.Count; i++)
            {
                var file = httpFileCollection[i];
                var stream = file.InputStream;
                if (stream == null || !stream.CanRead || stream.Length <= 0) continue;

                var workbook = WorkbookFactory.Create(stream);
                var tables = new List<DataTable>();
                for (int j = 0; j < workbook.NumberOfSheets; j++)
                {
                    var sheet = workbook.GetSheetAt(j);
                    if (sheet == null) continue;

                    var dataTable = ReadSheetToDataTable(sheet, headRowIndex, addEmptyRow);
                    tables.Add(dataTable);
                }
                collection.Add(tables);
            }

            return collection;
        }
        #endregion

        #region 读取 Excel 工作簿到 DataTable
        /// <summary>
        /// 读取 Excel 工作簿到 DataTable
        /// </summary>
        /// <param name="sheet">指定的 Sheet 工作簿</param>
        /// <param name="headRowIndex">表头行索引，默认：0，第一行</param>
        /// <param name="addEmptyRow">是否添加空行，默认为 false，不添加</param>
        /// <returns></returns>
        private DataTable ReadSheetToDataTable(ISheet sheet, int headRowIndex = 0, bool addEmptyRow = false)
        {
            var table = new DataTable(sheet.SheetName);
            var headerRow = sheet.GetRow(headRowIndex);
            if (headerRow == null || headerRow.LastCellNum < 0) throw new Exception("未获取到表头数据");
            var cellCount = headerRow.LastCellNum;
            var rowCount = sheet.LastRowNum;
            var startRowIndex = sheet.FirstRowNum;
            var RowIndex = 0;//Excel行号

            if (headRowIndex >= 0)
            {
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    //var column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                    //table.Columns.Add(column);
                    ICell cell = headerRow.GetCell(i);
                    if (cell != null)
                    {
                        string cellValue = GetCellValue(cell);
                        if (cellValue != null)
                        {
                            cellValue = cellValue.Trim().Replace(" ", "");
                            if (table.Columns.Contains(cellValue)) throw new Exception($"表头存在多个列名为“{cellValue}”的列！");
                            DataColumn column = new DataColumn(cellValue);
                            table.Columns.Add(column);
                        }
                        else
                        {
                            DataColumn column = new DataColumn("Column" + (i + 1));
                            table.Columns.Add(column);
                        }
                    }
                    else
                    {
                        DataColumn column = new DataColumn("Column" + (i + 1));
                        table.Columns.Add(column);
                    }
                }
                if (cellCount > 0)
                {
                    DataColumn column = new DataColumn("RowNum");//Excel行号
                    table.Columns.Add(column);
                }
                startRowIndex = headRowIndex + 1;
            }
            for (var i = startRowIndex; i <= rowCount; i++)
            {
                var dataRow = table.NewRow();
                var row = sheet.GetRow(i);
                if (row != null)
                {
                    RowIndex = row.RowNum + 1;
                    dataRow["RowNum"] = RowIndex;
                }
                else {
                    RowIndex++;
                    dataRow["RowNum"] = RowIndex;
                }
                if (row == null)//筛选空行
                {
                    if (addEmptyRow) table.Rows.Add(dataRow);
                    continue;
                }
                string blankStr = string.Empty;//空白单元格数
                for (var j = row.FirstCellNum; j < cellCount; j++)
                {
                    var cell = row.GetCell(j);
                    //判断单元格是否为空白
                    if (cell != null)
                    {
                        string cellVal = GetCellValue(cell);
                        dataRow[j] = cellVal.Trim();
                        if (string.IsNullOrEmpty(blankStr))blankStr = cellVal.Trim();
                    }
                }
                //筛选整个空单元格行
                if (string.IsNullOrEmpty(blankStr)) { if (addEmptyRow) table.Rows.Add(dataRow); continue; }
                if (row != null)table.Rows.Add(dataRow);
            }

            return table;
        }

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="cell">要获取值的 <see cref="ICell"/></param>
        /// <returns>从 <see cref="ICell"/> 获取到的值</returns>
        private string GetCellValue(ICell cell)
        {
            if (cell == null) return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: return string.Empty;
                case CellType.Boolean: return cell.BooleanCellValue.ToString();
                case CellType.Error: return cell.ErrorCellValue.ToString();
                case CellType.Numeric:
                    short format = cell.CellStyle.DataFormat;//是否是带格式的日期类型
                    if (format == 0xe || format == 0x16) { return cell.DateCellValue.ToString(); }
                    else { return cell.NumericCellValue.ToString(); }
                case CellType.Unknown:
                default: return cell.ToString();
                case CellType.String: return cell.StringCellValue;
                case CellType.Formula://公式
                    try
                    {
                        CellType CellType = cell.CachedFormulaResultType;
                        //判断是否公式计算值
                        if (CellType == CellType.String) return cell.StringCellValue;
                        else if (CellType == CellType.Numeric)
                        {
                            short format2 = cell.CellStyle.DataFormat;//是否是带格式的日期类型
                            if (format2 == 0xe || format2 == 0x16) { return cell.DateCellValue.ToString(); }
                            else { return cell.NumericCellValue.ToString(); }
                        }
                        else if (CellType == CellType.Blank) return string.Empty;
                        else if (CellType == CellType.Boolean) return cell.BooleanCellValue.ToString();
                        else
                        {
                            HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                            e.EvaluateInCell(cell);
                            return cell.ToString();
                        }
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }
        //下面是从官网扒来的cell类型匹配表：
        //0, "General"   //常规
        //1, "0"
        //2, "0.00"
        //3, "#,##0"
        //4, "#,##0.00"
        //5, "$#,##0_);($#,##0)"
        //6, "$#,##0_);[Red]($#,##0)"
        //7, "$#,##0.00);($#,##0.00)"
        //8, "$#,##0.00_);[Red]($#,##0.00)"
        //9, "0%"
        //0xa, "0.00%"
        //0xb, "0.00E+00"
        //0xc, "# ?/?"
        //0xd, "# ??/??"
        //0xe, "m/d/yy"
        //0xf, "d-mmm-yy"
        //0x10, "d-mmm"
        //0x11, "mmm-yy"
        //0x12, "h:mm AM/PM"
        //0x13, "h:mm:ss AM/PM"
        //0x14, "h:mm"
        //0x15, "h:mm:ss"
        //0x16, "m/d/yy h:mm"

        //// 0x17 - 0x24 reserved for international and undocumented 0x25, "#,##0_);(#,##0)"
        //0x26, "#,##0_);[Red](#,##0)"
        //0x27, "#,##0.00_);(#,##0.00)"
        //0x28, "#,##0.00_);[Red](#,##0.00)"
        //0x29, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"
        //0x2a, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"
        //0x2b, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
        //0x2c, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"
        //0x2d, "mm:ss"
        //0x2e, "[h]:mm:ss"
        //0x2f, "mm:ss.0"
        //0x30, "##0.0E+0"
        //0x31, "@" - This is text format.
        //0x31 "text" - Alias for "@"

        #endregion
    }

    /// <summary>
    /// 自定义单元格样式
    /// </summary>
    public class NPOIFontStyle
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public NPOIFontStyle() { }
        /// <param name="fontName">字体，默认黑体（黑体、宋体、Microsoft YaHei UI等Excel上所有字体）</param>
        /// <param name="fontWeight">字体粗细</param>
        /// <param name="fontColor">字体颜色<see cref="FontColor"/></param>
        /// <param name="verticalAlign">垂直对齐方式</param>
        public NPOIFontStyle(string fontName, short fontWeight, short fontColor)
        {
            this.fontName = fontName;
            this.fontWeight = fontWeight;
            this.fontColor = fontColor;
        }

        /// <summary>
        /// 字体,默认黑体（黑体、宋体、Microsoft YaHei UI等Excel上所有字体）
        /// </summary>
        public string fontName { get; set; }
        /// <summary>
        /// 字体粗细
        /// </summary>
        public short fontWeight { get; set; }
        /// <summary>
        /// 字体行高
        /// </summary>
        public double fontHeight { get; set; }
        /// <summary>
        /// 字号大小
        /// </summary>
        public double fontHeightInPoints { get; set; }
        /// <summary>
        /// 字体颜色<see cref="FontColor"/>
        /// </summary>
        public short fontColor { get; set; }
    }

    /// <summary>
    /// Excel导出拓展
    /// </summary>
    public class ExcelExport
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelExport() { }

        #region 属性

        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }
        /// <summary>
        /// 单元格索引
        /// </summary>
        public int CellIndex { get; set; }
        /// <summary>
        /// 当前表格行
        /// </summary>
        public IRow CurrentRow { get; set; }
        /// <summary>
        /// 当前表
        /// </summary>
        public ISheet CurrentSheet { get; set; }
        /// <summary>
        /// 当前工作簿
        /// </summary>
        public HSSFWorkbook CurrentBookXls { get; set; }
        /// <summary>
        /// 当前工作簿
        /// </summary>
        public XSSFWorkbook CurrentBookXlsx { get; set; }

        #endregion

        #region 封装方法

        /// <summary>
        /// 创建工作簿，创建的文件后缀必须为.xls
        /// </summary>
        /// <param name="book">工作簿</param>
        /// <param name="SheelName">表名</param>
        /// <returns></returns>
        public HSSFWorkbook CreateBookXls()
        {
            CurrentBookXls = new NPOI.HSSF.UserModel.HSSFWorkbook();
            return CurrentBookXls;
        }
        /// <summary>
        /// 创建工作簿，创建的文件后缀必须为.xlsx
        /// </summary>
        /// <param name="book">工作簿</param>
        /// <param name="SheelName">表名</param>
        /// <returns></returns>
        public XSSFWorkbook CreateBookXlsx()
        {
            CurrentBookXlsx = new XSSFWorkbook();
            return CurrentBookXlsx;
        }

        /// <summary>
        /// 创建表,可省略，创建Tr时会自动创建一个Sheet
        /// </summary>
        /// <param name="SheelName">表名</param>
        /// <returns></returns>
        public ISheet CreateSheet(string SheelName = "Sheet1")
        {
            if (CurrentBookXlsx != null) return CreateSheet(CurrentBookXlsx, SheelName);
            return CreateSheet(CurrentBookXls, SheelName);
        }
        /// <summary>
        /// 创建表
        /// </summary>
        /// <param name="book">工作簿</param>
        /// <param name="SheelName">表名</param>
        /// <returns></returns>
        public ISheet CreateSheet(HSSFWorkbook book, string SheelName = "Sheet1")
        {
            CurrentSheet = book.CreateSheet(SheelName);
            return CurrentSheet;
        }
        /// <summary>
        /// 创建表
        /// </summary>
        /// <param name="book">工作簿</param>
        /// <param name="SheelName">表名</param>
        /// <returns></returns>
        public ISheet CreateSheet(XSSFWorkbook book, string SheelName = "Sheet1")
        {
            CurrentSheet = book.CreateSheet(SheelName);
            return CurrentSheet;
        }
        /// <summary>
        /// 合并单元格，例:合并第2行至第3行的第5列至第8列，应为（1,2,4,7）,
        /// </summary>
        /// <param name="firstRow">合并起始行索引</param>
        /// <param name="lastRow">合并结束行索引</param>
        /// <param name="firstColumn">起始列索引</param>
        /// <param name="lastColumn">结束列索引</param>
        public void MergeCell(int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            CurrentSheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow, firstColumn, lastColumn));
        }

        /// <summary>
        /// 创建表头，常规不加粗字体，与内容字体一致
        /// </summary>
        /// <param name="columns">表头列</param>
        public IRow CreateHead(string[] columns)
        {
            RowIndex = 0;
            IRow headRow = CreateTr(0);
            foreach (var item in columns) CreateTd(item);
            return headRow;
        }
        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="columns">表头列</param>
        /// <param name="style">单元格样式</param>
        public IRow CreateHead(string[] columns, ICellStyle style)
        {
            RowIndex = 0;
            IRow headRow = CreateTr(0);
            foreach (var item in columns) CreateTd(item, style);
            return headRow;
        }
        /// <summary>
        /// 创建表头并添加默认样式：黑体、加粗
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        public IRow CreateHeadDefaultStyle(string[] columns)
        {
            if (CurrentBookXls != null)
            {
                //单元格样式
                ICellStyle HeadStyle = CreateStyle(new NPOIFontStyle("黑体", 700, NPOI.HSSF.Util.HSSFColor.Black.Index));
                return CreateHead(columns, HeadStyle);
            }
            else
            {
                //单元格样式
                ICellStyle HeadStyle = CreateStyle(new NPOIFontStyle("黑体", 700, NPOI.HSSF.Util.HSSFColor.Black.Index));
                return CreateHead(columns, HeadStyle);
            }

        }
        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="sheet">工作簿</param>
        /// <param name="rownum">行号</param>
        /// <param name="columns">表头列</param>
        /// <param name="style">单元格样式</param>
        public IRow CreateHead(ISheet sheet, int rownum, string[] columns)
        {
            IRow headRow = CreateTr(sheet, rownum);
            CellIndex = 0;
            foreach (var item in columns) CreateTd(item);
            return headRow;
        }
        /// <summary>
        /// 创建表头
        /// </summary>
        /// <param name="sheet">工作簿</param>
        /// <param name="rownum">行号</param>
        /// <param name="columns">表头列</param>
        /// <param name="style">单元格样式</param>
        public IRow CreateHead(ISheet sheet, int rownum, string[] columns, ICellStyle style)
        {
            IRow headRow = CreateTr(sheet, rownum);
            CellIndex = 0;
            foreach (var item in columns) CreateTd(item);
            return headRow;
        }

        /// <summary>
        /// 创建行
        /// </summary>
        /// <param name="rownum">自定义传入行号，默认：-1，自动累加行号</param>
        /// <returns></returns>
        public IRow CreateTr(int rownum = -1)
        {
            CellIndex = 0;
            if (rownum == -1) ++RowIndex;
            if (rownum >= 0) this.RowIndex = rownum;
            if (CurrentSheet == null) CreateSheet();//自动添加一个Sheet表格
            CurrentRow = CurrentSheet.CreateRow(RowIndex);
            return CurrentRow;
        }
        /// <summary>
        /// 创建行
        /// </summary>
        /// <param name="sheet">绘制表</param>
        /// <param name="rownum">自定义传入行号，默认：-1，自动累加行号</param>
        /// <returns></returns>
        public IRow CreateTr(ISheet sheet, int rownum = -1)
        {
            CellIndex = 0;
            if (rownum == -1) ++RowIndex;
            if (rownum >= 0) this.RowIndex = rownum;
            this.CurrentSheet = sheet;
            CurrentRow = CurrentSheet.CreateRow(RowIndex);
            return CurrentRow;
        }

        /// <summary>
        /// 创建单元格样式
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        public ICellStyle CreateStyle()
        {
            if (CurrentBookXls != null)
            {
                //单元格样式
                ICellStyle _style = CurrentBookXls.CreateCellStyle();
                return _style;
            }
            else
            {
                //单元格样式
                ICellStyle _style = CurrentBookXls.CreateCellStyle();
                return _style;
            }

        }
        /// <summary>
        /// 创建单元格样式：黑体、加粗
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        public ICellStyle CreateStyle(NPOIFontStyle style)
        {
            if (CurrentBookXls != null)
            {
                //单元格样式
                ICellStyle _style = CreateStyle();
                HSSFFont font = (HSSFFont)CurrentBookXls.CreateFont();
                if (string.IsNullOrEmpty(style.fontName)) font.FontName = style.fontName;//字体
                if (style.fontWeight > 0) font.Boldweight = style.fontWeight;//加粗细
                if (style.fontColor > 0) font.Color = style.fontColor;//字体颜色
                if (style.fontHeight > 0) font.FontHeight = style.fontHeight;//字体行高
                if (style.fontHeightInPoints > 0) font.FontHeight = style.fontHeightInPoints;//字号大小
                _style.SetFont(font);
                return _style;
            }
            else
            {
                //单元格样式
                ICellStyle _style = CurrentBookXls.CreateCellStyle();
                HSSFFont font = (HSSFFont)CurrentBookXls.CreateFont();
                if (string.IsNullOrEmpty(style.fontName)) font.FontName = style.fontName;//字体
                if (style.fontWeight > 0) font.Boldweight = style.fontWeight;//加粗细
                if (style.fontColor > 0) font.Color = style.fontColor;//字体颜色
                if (style.fontHeight > 0) font.FontHeight = style.fontHeight;//字体行高
                if (style.fontHeightInPoints > 0) font.FontHeight = style.fontHeightInPoints;//字号大小
                _style.SetFont(font);
                return _style;
            }

        }

        /// <summary>
        /// 创建字体样式
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        public HSSFFont CreateFont()
        {
            if (CurrentBookXls != null)
            {
                HSSFFont font = (HSSFFont)CurrentBookXls.CreateFont();
                return font;
            }
            else
            {
                HSSFFont font = (HSSFFont)CurrentBookXls.CreateFont();
                return font;
            }
        }
        #endregion

        #region CreateTd 单元格方法重载
        /// <summary>
        /// 创建一个空白单元格
        /// </summary>
        public ICell CreateTd() => CreateTd(string.Empty);
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <returns></returns>
        public ICell CreateTd(string Value)
        {
            try
            {
                ICell cell = CurrentRow.CreateCell(CellIndex);
                if (!string.IsNullOrEmpty(Value)) cell.SetCellValue(Value.ToString());
                CellIndex++;
                return cell;
            }
            catch (Exception e)
            {

                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <returns></returns>
        public ICell CreateTd(int? Value)
        {
            if (Value != null) return this.CreateTd(Value.ToString());
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <returns></returns>
        public ICell CreateTd(decimal? Value)
        {
            if (Value != null) return this.CreateTd(Value.ToString());
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public ICell CreateTd(decimal? Value, string Format)
        {
            if (Value != null) return this.CreateTd(Value.Value.ToString(Format));
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <returns></returns>
        public ICell CreateTd(double? Value)
        {
            if (Value != null) return this.CreateTd(Value.ToString());
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public ICell CreateTd(double? Value, string Format)
        {
            if (Value != null) return this.CreateTd(Value.Value.ToString(Format));
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <returns></returns>
        public ICell CreateTd(float? Value)
        {
            if (Value != null) return this.CreateTd(Value.ToString());
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Format">保留格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public ICell CreateTd(float? Value, string Format)
        {
            if (Value != null) return this.CreateTd(Value.Value.ToString(Format));
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <returns></returns>
        public ICell CreateTd(bool? Value)
        {
            if (Value != null) return this.CreateTd(Value.ToString());
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Format">时间格式</param>
        /// <returns>返回<see cref="ICell"/>单元格</returns>
        public ICell CreateTd(DateTime? Value, string Format = "yyyy-MM-dd")
        {
            if (Value != null) return this.CreateTd(Value.Value.ToString(Format));
            else return this.CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns>返回<see cref="int"/>列索引</returns>
        public ICell CreateTd(string Value, ICellStyle Style)
        {
            try
            {
                ICell cell = CurrentRow.CreateCell(CellIndex);
                if (!string.IsNullOrEmpty(Value)) cell.SetCellValue(Value);
                if (Style != null) cell.CellStyle = Style;
                CellIndex++;
                return cell;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns></returns>
        public ICell CreateTd(int? Value, ICellStyle Style)
        {
            if (Value != null) return CreateTd(Value.ToString(), Style);
            else return CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns></returns>
        public ICell CreateTd(decimal? Value, ICellStyle Style)
        {
            if (Value != null) return CreateTd(Value.ToString(), Style);
            else return CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns></returns>
        public ICell CreateTd(decimal? Value, string Format, ICellStyle Style)
        {
            if (Value != null) return CreateTd(Value.Value.ToString(Format), Style);
            else return CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns></returns>
        public ICell CreateTd(double? Value, ICellStyle Style)
        {
            if (Value != null) return CreateTd(Value.ToString(), Style);
            else return CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns></returns>
        public ICell CreateTd(double? Value, string Format, ICellStyle Style)
        {
            if (Value != null) return CreateTd(Value.Value.ToString(Format), Style);
            else return CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns></returns>
        public ICell CreateTd(float? Value, ICellStyle Style)
        {
            if (Value != null) return this.CreateTd(Value.ToString(), Style);
            else return CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns></returns>
        public ICell CreateTd(float? Value, string Format, ICellStyle Style)
        {
            if (Value != null) return CreateTd(Value.Value.ToString(Format), Style);
            else return CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <returns></returns>
        public ICell CreateTd(bool? Value, ICellStyle Style)
        {
            if (Value != null) return CreateTd(Value.ToString(), Style);
            else return CreateTd(string.Empty);
        }
        /// <summary>
        /// 创建一个单元格
        /// </summary>
        /// <param name="Value">值</param>
        /// <param name="Style">单元格样式</param>
        /// <param name="Format">保留格式</param>
        /// <returns></returns>
        public ICell CreateTd(DateTime? Value, ICellStyle Style, string Format = "yyyy-MM-dd")
        {
            if (Value != null) return CreateTd(Value.Value.ToString(Format), Style);
            else return CreateTd(string.Empty);
        }

        #endregion

        #region 表格自动列宽

        /// <summary>
        /// 自动设置列宽，放在内容循环值后
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columns"></param>
        public void AutomaticSetWidth(ISheet sheet = null, int columnCount = -1)
        {
            //获取当前列数
            if (columnCount == -1) columnCount = CellIndex;
            //自动列宽
            if (sheet != null) for (int i = 0; i < columnCount; i++) sheet.AutoSizeColumn(i, true);
            else { for (int i = 0; i < columnCount; i++) CurrentSheet.AutoSizeColumn(i, true); }
        }

        #endregion

        #region 保存文件到服务器

        /// <summary>
        /// 保存要导出的文件,返回服务器存储的文件相对地址：Resource\Export\Excel\文件名.xls
        /// </summary>
        /// <param name="sFileName">Excel文件名</param>
        /// <param name="filePath">保存绝对路径,如：D:\\User\\Projects\\Resource\\Export\\Excel\\</param>
        public string ExportToExcel(string sFileName, string filePath)
        {
            sFileName = string.Format("{0}_{1}", Guid.NewGuid().ToString().Replace("-", string.Empty).ToLower(), sFileName);
            string sRoot = filePath;// GlobalContext.HostingEnvironment.ContentRootPath;
            string partDirectory = string.Format("Resource{0}Export{0}Excel", Path.DirectorySeparatorChar);
            string sDirectory = Path.Combine(sRoot, partDirectory);
            string sFilePath = Path.Combine(sDirectory, sFileName);
            if (!Directory.Exists(sDirectory))
            {
                Directory.CreateDirectory(sDirectory);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                if (CurrentBookXls != null) CurrentBookXls.Write(ms);
                else if (CurrentBookXlsx != null) CurrentBookXlsx.Write(ms);
                else throw new Exception("未获取到工作簿");
                ms.Flush();
                ms.Position = 0;
                using (FileStream fs = new FileStream(sFilePath, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
            return partDirectory + Path.DirectorySeparatorChar + sFileName;
        }
        /// <summary>
        /// 保存要导出的文件,返回服务器存储的文件地址
        /// </summary>
        /// <param name="sFileName">Excel文件名</param>
        /// <param name="filePath">保存绝对路径,如：D:\\User\\Projects\\Resource\\Export\\Excel\\</param>
        /// <param name="book">工作簿</param>
        public string ExportToExcel(string sFileName, string filePath, HSSFWorkbook book)
        {
            sFileName = string.Format("{0}_{1}", Guid.NewGuid().ToString().Replace("-", string.Empty).ToLower(), sFileName);
            string sRoot = filePath;// GlobalContext.HostingEnvironment.ContentRootPath;
            string partDirectory = string.Format("Resource{0}Export{0}Excel", Path.DirectorySeparatorChar);
            string sDirectory = Path.Combine(sRoot, partDirectory);
            string sFilePath = Path.Combine(sDirectory, sFileName);
            if (!Directory.Exists(sDirectory))
            {
                Directory.CreateDirectory(sDirectory);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                book.Write(ms);
                ms.Flush();
                ms.Position = 0;
                using (FileStream fs = new FileStream(sFilePath, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
            return partDirectory + Path.DirectorySeparatorChar + sFileName;
        }
        /// <summary>
        /// 保存要导出的文件,返回服务器存储的文件地址
        /// </summary>
        /// <param name="sFileName">Excel文件名</param>
        /// <param name="filePath">保存绝对路径,如：D:\\User\\Projects\\Resource\\Export\\Excel\\</param>
        /// <param name="book">工作簿</param>
        public string ExportToExcel(string sFileName, string filePath, XSSFWorkbook book)
        {
            sFileName = string.Format("{0}_{1}", Guid.NewGuid().ToString().Replace("-", string.Empty).ToLower(), sFileName);
            string sRoot = filePath;// GlobalContext.HostingEnvironment.ContentRootPath;
            string partDirectory = string.Format("Resource{0}Export{0}Excel", Path.DirectorySeparatorChar);
            string sDirectory = Path.Combine(sRoot, partDirectory);
            string sFilePath = Path.Combine(sDirectory, sFileName);
            if (!Directory.Exists(sDirectory))
            {
                Directory.CreateDirectory(sDirectory);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                book.Write(ms);
                ms.Flush();
                ms.Position = 0;
                using (FileStream fs = new FileStream(sFilePath, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fs.Write(data, 0, data.Length);
                    fs.Flush();
                }
            }
            return partDirectory + Path.DirectorySeparatorChar + sFileName;
        }

        /// <summary>
        /// 将创建的工作簿输出到二进制流
        /// </summary>
        public byte[] ToByteArray()
        {
            using (MemoryStream ms = new MemoryStream())
            {
                if (CurrentBookXls != null) CurrentBookXls.Write(ms);
                else if (CurrentBookXlsx != null) CurrentBookXlsx.Write(ms);
                else throw new Exception("未获取到工作簿");
                ms.Flush();
                ms.Position = 0;
                byte[] value = ms.ToArray();
                if (CurrentBookXls != null) CurrentBookXls.Close();
                else if (CurrentBookXlsx != null) CurrentBookXlsx.Close();
                return value;
            }
        }
        #endregion
    }
}