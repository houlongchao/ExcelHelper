using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelper.Aspose
{
    /// <summary>
    /// Aspose Excel Helper
    /// </summary>
    public static class AsposeCellHelper
    {
        #region File to Workbook

        /// <summary>
        /// 读取 Excel
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static Workbook ReadExcel(string filePath)
        {
            var fileBytes = File.ReadAllBytes(filePath);
            using (var stream = new MemoryStream(fileBytes))
            {
                stream.Position = 0;
                return new Workbook(stream);
            }
        }

        /// <summary>
        /// 读取 Excel
        /// </summary>
        /// <param name="fileBytes"></param>
        /// <returns></returns>
        public static Workbook ReadExcel(byte[] fileBytes)
        {
            using (var stream = new MemoryStream(fileBytes))
            {
                return new Workbook(stream);
            }
        }

        /// <summary>
        /// 读取 Excel
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static Workbook ReadExcel(Stream stream)
        {
            stream.Position = 0;
            return new Workbook(stream);
        }

        /// <summary>
        /// 创建 Excel
        /// </summary>
        /// <returns></returns>
        public static Workbook CreateExcel()
        {
            return new Workbook();
        }

        #endregion

        #region File to Sheet

        /// <summary>
        /// 读取指定 Sheet
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static Worksheet ReadExcelSheet(string filePath, int index)
        {
            return ReadExcel(filePath).GetSheetAt(index);
        }

        /// <summary>
        /// 读取指定 Sheet
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Worksheet ReadExcelSheet(string filePath, string name)
        {
            return ReadExcel(filePath).GetSheet(name);
        }

        /// <summary>
        /// 读取指定 Sheet
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Worksheet ReadExcelSheet(Stream stream, string name)
        {
            return ReadExcel(stream).GetSheet(name);
        }

        #endregion

        #region Workbook Extensions

        /// <summary>
        /// 转换为 byte 数据
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static byte[] ToBytes(this Workbook workbook)
        {
            using (var stream = workbook.SaveToStream())
            {
                return stream.ToArray();
            }
        }

        /// <summary>
        /// 写入到文件
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fileName"></param>
        public static void ToFile(this Workbook workbook, string fileName)
        {
            workbook.Save(fileName);
        }

        /// <summary>
        /// 计算 Excel
        /// </summary>
        /// <param name="workbook"></param>
        public static void EvaluateAllFormulaCells(this Workbook workbook)
        {
            workbook.CalculateFormula(true);
        }

        /// <summary>
        /// 获取指定 Sheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Worksheet GetSheet(this Workbook workbook, string name)
        {
            return workbook.Worksheets[name];
        }

        /// <summary>
        /// 获取指定 Sheet，如果指定多个 name ， 则依次获取，获取到第一个匹配的 Sheet 返回
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="names"></param>
        /// <returns></returns>
        public static Worksheet GetSheet(this Workbook workbook, params string[] names)
        {
            foreach (var name in names)
            {
                var sheet = workbook.Worksheets[name];
                if (sheet != null)
                {
                    return sheet;
                }
            }
            return null;
        }

        /// <summary>
        /// 获取指定 Sheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static Worksheet GetSheetAt(this Workbook workbook, int index)
        {
            return workbook.Worksheets[index];
        }

        /// <summary>
        /// 创建 Sheet
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Worksheet CreateSheet(this Workbook workbook, string name)
        {
            return workbook.Worksheets.Add(name);
        }

        #endregion

        #region Sheet Extensions

        /// <summary>
        /// 获取指定行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static Row GetRow(this Worksheet sheet, int row)
        {
            return sheet.Cells.GetRow(row);
        }

        /// <summary>
        /// 创建指定行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static Row CreateRow(this Worksheet sheet, int row)
        {
            sheet.Cells.InsertRow(row);
            return sheet.GetRow(row);
        }

        /// <summary>
        /// 获取指定单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static Cell GetCell(this Worksheet sheet, int row, int cell)
        {
            return sheet.Cells[row, cell];
        }

        /// <summary>
        /// 获取指定单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <returns></returns>
        public static Cell GetCell(this Worksheet sheet, string cellRef)
        {
            return sheet.Cells[cellRef];
        }

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <returns></returns>
        public static Cell CreateCell(this Worksheet sheet, string cellRef)
        {
            return sheet.Cells[cellRef];
        }

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static Cell CreateCell(this Worksheet sheet, int row, int col)
        {
            return sheet.Cells[row, col];
        }

        /// <summary>
        /// 获取或创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static Cell GetOrCreateCell(this Worksheet sheet, int row, int cell)
        {
            return sheet.GetCell(row, cell) ?? sheet.CreateCell(row, cell);
        }

        /// <summary>
        /// 获取或创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <returns></returns>
        public static Cell GetOrCreateCell(this Worksheet sheet, string cellRef)
        {
            return sheet.GetCell(cellRef) ?? sheet.CreateCell(cellRef);
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetCell(this Worksheet sheet, string cellRef, bool data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.Value = data;
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetCell(this Worksheet sheet, string cellRef, double data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.Value = data;
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetCell(this Worksheet sheet, string cellRef, string data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.Value = data;
            return cell;
        }


        /// <summary>
        /// 获取Sheet的总行数
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static int GetRowCount(this Worksheet sheet)
        {
            return sheet.Cells.MaxRow + 1;
        }

        #endregion

        #region Row Extensions

        /// <summary>
        /// 获取指定单元格
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static Cell GetCell(this Row row, int cell)
        {
            return row.GetCellOrNull(cell);
        }

        /// <summary>
        /// 获取指定标题<paramref name="text"/>的列 Index
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="text">要匹配的列内容</param>
        /// <param name="defaultIndex">如果没有匹配默认返回的Index</param>
        /// <param name="otherTexts">除过<paramref name="text"/>的其它匹配内容</param>
        /// <returns></returns>
        public static int GetIndex(this Row row, string text, int defaultIndex = -1, params string[] otherTexts)
        {
            for (int i = 0; i < row.LastCell.Column; i++)
            {
                var cellValue = row.GetCellOrNull(i)?.StringValue;
                if (cellValue == text)
                {
                    return i;
                }
                foreach (var otherText in otherTexts)
                {
                    if (cellValue == otherText)
                    {
                        return i;
                    }
                }
            }
            return defaultIndex;
        }


        #endregion

        #region Cell Extensions

        /// <summary>
        /// 获取数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="calculate"></param>
        /// <returns></returns>
        public static object GetData(this Cell cell, bool calculate = true)
        {
            if (cell == null)
            {
                return null;
            }
            if (calculate)
            {
                return cell.Value;
            }
            return cell.Formula;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetValue(this Cell cell, string data)
        {
            cell.Value = data;
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetValue(this Cell cell, double data)
        {
            cell.Value = data;
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetValue(this Cell cell, DateTime data)
        {
            cell.Value = data;
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetValue(this Cell cell, bool data)
        {
            cell.Value = data;
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetValue(this Cell cell, object data)
        {
            if (data is DateTime dt)
            {
                if (DateTime.MinValue != dt)
                {
                    cell.SetValue(dt);
                }
            }
            else if (data is bool b)
            {
                cell.SetValue(b);
            }
            else if (data is double d)
            {
                cell.SetValue(d);
            }
            else if (data is int di)
            {
                cell.SetValue(di);
            }
            else if (data is decimal dc)
            {
                cell.SetValue((double)dc);
            }
            else
            {
                cell.SetValue(data?.ToString());
            }
            return cell;
        }

        /// <summary>
        /// 设置单元格格式字符串
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="format"></param>
        /// <returns></returns>
        public static Cell SetDataFormat(this Cell cell, string format = "yyyy-MM-dd")
        {
            var cellStyle = cell.Worksheet.Workbook.CreateStyle();

            cellStyle.Copy(cell.GetStyle());
            cellStyle.Custom = format;
            cell.SetStyle(cellStyle);

            return cell;
        }

        /// <summary>
        /// 设置单元格备注
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static Cell SetComment(this Cell cell, string data)
        {
            var index = cell.Worksheet.Comments.Add(cell.Name);
            var comment = cell.Worksheet.Comments[index];
            comment.Note = data;
            return cell;
        }

        /// <summary>
        /// 设置单元格边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="borderStype"></param>
        /// <returns></returns>
        public static Cell SetBorder(this Cell cell, CellBorderType borderStype = CellBorderType.Thin)
        {
            return cell.SetBorder(borderStype, borderStype, borderStype, borderStype);
        }

        /// <summary>
        /// 设置单元格边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="borderTop"></param>
        /// <param name="borderRight"></param>
        /// <param name="borderBottom"></param>
        /// <param name="borderLeft"></param>
        /// <returns></returns>
        public static Cell SetBorder(this Cell cell, CellBorderType borderTop, CellBorderType borderRight, CellBorderType borderBottom, CellBorderType borderLeft)
        {
            var cellStyle = cell.Worksheet.Workbook.CreateStyle();

            cellStyle.Copy(cell.GetStyle());

            cellStyle.Borders[BorderType.TopBorder].LineStyle = borderTop;
            cellStyle.Borders[BorderType.RightBorder].LineStyle = borderRight;
            cellStyle.Borders[BorderType.BottomBorder].LineStyle = borderBottom;
            cellStyle.Borders[BorderType.LeftBorder].LineStyle = borderLeft;
            cell.SetStyle(cellStyle);
            return cell;
        }

        /// <summary>
        /// 设置字体
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="fontAction"></param>
        /// <returns></returns>
        public static Cell SetFont(this Cell cell, Action<Font> fontAction)
        {
            var cellStyle = cell.Worksheet.Workbook.CreateStyle();
            cellStyle.Copy(cell.GetStyle());

            fontAction?.Invoke(cellStyle.Font);

            cell.SetStyle(cellStyle);

            return cell;
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public static Cell SetCellStyle(this Cell cell, Style cellStyle)
        {
            cell.SetStyle(cellStyle);
            return cell;
        }

        /// <summary>
        /// 设置图片
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static Cell SetImage(this Cell cell, byte[] bytes)
        {
            cell.Worksheet.Pictures.Add(cell.Row, cell.Column, cell.Row + 1, cell.Column + 1, new MemoryStream(bytes));
            return cell;
        }

        /// <summary>
        /// 获取单元格图片数据
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static byte[] GetImage(this Cell cell)
        {
            var pictureData = GetPictureData(cell.Worksheet);
            if (pictureData.TryGetValue(cell, out var value))
            {
                return value;
            }
            return null;
        }


        #endregion


        #region Sheet Image

        /// <summary>
        /// 获取图片字典
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static Dictionary<Cell, byte[]> GetPictureData(this Worksheet sheet)
        {
            var result = new Dictionary<Cell, byte[]>();

            foreach (var picture in sheet.Pictures)
            {
                var cell = sheet.Cells[picture.UpperLeftRow, picture.UpperLeftColumn];
                result[cell] = picture.Data;
            }

            return result;
        }

        #endregion

    }
}
