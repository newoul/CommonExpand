using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using DataTable = System.Data.DataTable;

namespace ExpandComponents
{
    /********************************************************************************

    ** 类名称： ExcelExpand

    ** 描述：导入导出Excel拓展

    ** 引用： NPOI.dll   NPOI.OOML.dll  NPOI.OpenXml4Net.dll  NPOI.OpenXmlFormats.dll  ICSharpCode.SharpZipLib.dll(版本：1.0.0.999)

    ** 作者： LW

    *********************************************************************************/

    /// <summary>
    /// Excel导入拓展
    /// </summary>
    public static class ExcelImport
    {
        #region 导入读取文件拓展

        #region NPOI 导入

        #region 将 Excel 文件读取到 DataTable
        /// <summary>
        /// 将 Excel 文件读取到 <see cref="DataTable"/>，多用于服务器读取Excel
        /// </summary>
        /// <param name="filePath">文件完整路径名,文件绝对路径，多用于服务器读取Excel</param>
        /// <param name="sheetName">指定读取 Excel 工作薄 sheet 的名称</param>
        /// <param name="firstRowIsColumnName">首行是否为 <see cref="DataColumn.ColumnName"/></param>
        /// <returns><see cref="DataTable"/>数据表</returns>
        public static DataTable ReadExcelToDataTable(string filePath, string sheetName = null, bool firstRowIsColumnName = true)
        {
            if (string.IsNullOrEmpty(filePath)) return null;
            if (!File.Exists(filePath)) return null;

            //根据指定路径读取文件
            var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            if (fileStream == null || fileStream.Length <= 0) return null;

            //定义要返回的datatable对象
            DataTable data = new DataTable();
            //Excel行号
            var RowIndex = 0;
            //excel工作表
            ISheet sheet = null;

            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
                //根据文件流创建excel数据结构
                IWorkbook workbook = WorkbookFactory.Create(fileStream);


                if (string.IsNullOrEmpty(sheetName)) sheet = workbook.GetSheetAt(0);
                else
                {
                    sheet = workbook.GetSheet(sheetName);

                    //如果没有找到指定的sheetName对应的sheet，则获取第一个sheet
                    if (sheet == null) sheet = workbook.GetSheetAt(0);
                }

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    if (firstRow == null) new Exception("未获取到表头数据");
                    RowIndex = firstRow.RowNum + 1;
                    //一行最后一个cell的编号 即总的列数
                    int cellCount = firstRow.LastCellNum;

                    //如果第一行是标题列名
                    if (firstRowIsColumnName)
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
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; // 没有数据的行默认是 null，如果为 null 则不添加
                        var blankCount = 0;//空白单元格数
                        DataRow dataRow = data.NewRow();
                        RowIndex = row.RowNum + 1;
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
                                    //判断是否日期类型
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim();
                                    }
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
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                //fileStream.Flush();
                fileStream.Close();
                fileStream.Dispose();
            }
        }
        /// <summary>
        /// 将 <see cref="Stream"/> 对象读取到 <see cref="DataTable"/>,指定表头索引，多用于服务器读取Excel
        /// </summary>
        /// <param name="filePath">文件完整路径名,文件绝对路径，多用于服务器读取Excel</param>
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

            //定义要返回的datatable对象
            DataTable data = new DataTable();
            //Excel行号
            var RowIndex = 1;
            //excel工作表
            ISheet sheet = null;

            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
                //根据文件流创建excel数据结构
                IWorkbook workbook = WorkbookFactory.Create(fileStream);


                if (string.IsNullOrEmpty(sheetName)) sheet = workbook.GetSheetAt(0);
                else
                {
                    sheet = workbook.GetSheet(sheetName);

                    //如果没有找到指定的sheetName对应的sheet，则获取第一个sheet
                    if (sheet == null) sheet = workbook.GetSheetAt(0);
                }

                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(HeaderIndex);
                    if (firstRow == null || firstRow.FirstCellNum < 0) new Exception("未获取到表头数据");
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
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null || row.FirstCellNum < 0) continue; // 没有数据的行默认是 null，如果为 null 则不添加
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
                                    //判断是否日期类型
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim();
                                    }
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
            catch (Exception e)
            {
                var err = new Exception($"行号:{RowIndex},错误：{e.Message}");
                throw err;
            }
            finally
            {
                fileStream.Close();
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
        public static DataTable ReadStreamToDataTable(Stream stream, string sheetName = null, bool firstRowIsColumnName = true)
        {
            if (stream == null || stream.Length <= 0) return null;

            //定义要返回的datatable对象
            var data = new DataTable();

            //excel工作表
            ISheet sheet = null;
            //Excel行号
            var RowIndex = 1;
            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
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
                    IRow firstRow = sheet.GetRow(0);
                    if (firstRow == null || firstRow.FirstCellNum < 0)throw new Exception("未获取到表头数据");
                    RowIndex = firstRow.RowNum + 1;
                    //一行最后一个cell的编号 即总的列数
                    int cellCount = firstRow.LastCellNum;

                    //如果第一行是标题列名
                    if (firstRowIsColumnName)
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
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
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
                                    //判断是否日期类型
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim();
                                    }
                                }
                                else if (cell.CellType == CellType.Formula)
                                {
                                    CellType CellType = row.GetCell(j).CachedFormulaResultType;
                                    //判断是否公式计算值
                                    if(CellType == CellType.String)
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

            //定义要返回的datatable对象
            var data = new DataTable();

            //excel工作表
            ISheet sheet = null;
            //Excel行号
            var RowIndex = 1;
            //数据开始行(排除标题行)
            int startRow = 0;
            try
            {
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
                    if (firstRow == null || firstRow.FirstCellNum < 0) new Exception("未获取到表头数据");
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
                                    //判断是否日期类型
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue;
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString().Trim();
                                    }
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
                        if (RowIndex >= 61)
                        {

                        }
                        if (cellCount > 0)//行号
                        {
                            var ColumnIndex = data.Columns.Count - 1;
                            dataRow[ColumnIndex] = RowIndex.ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
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

        #endregion

        /// <summary>
        /// 是否为Excel文件,是否为.xls，.xlsx两种类型的文件
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns></returns>
        public static bool IsExcelFile(string path)
        {
            var _fileInfo = new FileInfo(path);
            if (_fileInfo == null) return false;
            var ext = _fileInfo.Extension.ToLower();
            if (ext == ".xls" || ext == ".xlsx") return true;
            return false;
        }

        #endregion

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
        /// 创建表
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
        /// 创建表头
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
        /// <param name="rownum">行号</param>
        /// <returns></returns>
        public IRow CreateTr(int rownum = -1)
        {
            CellIndex = 0;
            if (rownum == -1) ++RowIndex;
            CurrentRow = CurrentSheet.CreateRow(RowIndex);
            return CurrentRow;
        }
        /// <summary>
        /// 创建行
        /// </summary>
        /// <param name="sheet">绘制表</param>
        /// <param name="rownum">行号</param>
        /// <returns></returns>
        public IRow CreateTr(ISheet sheet, int rownum = -1)
        {
            CellIndex = 0;
            if (rownum == -1) ++RowIndex;
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
        public string ExportToExcel(string sFileName,string filePath)
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