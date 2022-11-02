using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;

namespace ExcelHelper.NPOI
{
    /// <summary>
    /// NPOI 帮助类
    /// </summary>
    public static class NpoiHelper
    {
        #region File to Workbook

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static IWorkbook ReadExcel(string filePath)
        {
            var fileBytes = File.ReadAllBytes(filePath);
            using (var stream = new MemoryStream(fileBytes))
            {
                stream.Position = 0;
                if (filePath.EndsWith(".xlsx")) // 2007版本
                {
                    return new XSSFWorkbook(stream);
                }
                if (filePath.EndsWith(".xls")) // 2003版本
                {
                    return new HSSFWorkbook(stream);
                }
                return new XSSFWorkbook(stream);
            }
        }

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="fileBytes"></param>
        /// <returns></returns>
        public static IWorkbook ReadExcel(byte[] fileBytes)
        {
            using (var stream = new MemoryStream(fileBytes))
            {
                return new XSSFWorkbook(stream);
            }
        }

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static IWorkbook ReadExcel(Stream stream)
        {
            stream.Position = 0;
            return new XSSFWorkbook(stream);
        }

        /// <summary>
        /// 创建一个Excel操作对象
        /// </summary>
        /// <returns></returns>
        public static IWorkbook CreateExcel()
        {
            return new XSSFWorkbook();
        }

        #endregion

        #region File to Sheet

        /// <summary>
        /// 读取指定 Sheet
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static ISheet ReadExcelSheet(string filePath, int index)
        {
            return ReadExcel(filePath).GetSheetAt(index);
        }

        /// <summary>
        /// 读取指定 Sheet
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static ISheet ReadExcelSheet(string filePath, string name)
        {
            return ReadExcel(filePath).GetSheet(name);
        }

        /// <summary>
        /// 读取指定 Sheet
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static ISheet ReadExcelSheet(Stream stream, string name)
        {
            return ReadExcel(stream).GetSheet(name);
        }

        #endregion

        #region Workbook Extensions

        /// <summary>
        /// 将Excel操作对象转为二进制文件
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static byte[] ToBytes(this IWorkbook workbook)
        {
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// 计算所有公式
        /// </summary>
        /// <param name="workbook"></param>
        public static void EvaluateAllFormulaCells(this IWorkbook workbook)
        {
            BaseFormulaEvaluator.EvaluateAllFormulaCells(workbook);
        }

        /// <summary>
        /// 计算所有公式
        /// </summary>
        /// <param name="workbook"></param>
        public static void EvaluateAllFormulaCellsIgnoreError(this IWorkbook workbook)
        {
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                foreach (IRow item in workbook.GetSheetAt(i))
                {
                    foreach (var item2 in item)
                    {
                        if (item2.CellType == CellType.Formula)
                        {
                            try
                            {
                                evaluator.EvaluateFormulaCell(item2);
                            }
                            catch
                            { //不做处理，为了继续执行后续单元格 }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 获得公式计算器
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static IFormulaEvaluator CreateFormulaEvaluator(this IWorkbook workbook)
        {
            return workbook.GetCreationHelper().CreateFormulaEvaluator();
        }

        /// <summary>
        /// 获取指定Sheet页，可以指定多个Sheet名称依次匹配，无匹配项返回<c>null</c>
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="names"></param>
        /// <returns></returns>
        public static ISheet GetSheet(this IWorkbook workbook, params string[] names)
        {
            foreach (var name in names)
            {
                var sheet = workbook.GetSheet(name);
                if (sheet != null)
                {
                    return sheet;
                }
            }
            return null;
        }

        #endregion

        #region Sheet Extensions

        /// <summary>
        /// 读取指定单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static ICell GetCell(this ISheet sheet, int row, int cell)
        {
            var sheetRow = sheet.GetRow(row);
            return sheetRow?.GetCell(cell);
        }

        /// <summary>
        /// 读取指定单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <returns></returns>
        public static ICell GetCell(this ISheet sheet, string cellRef)
        {
            var cr = new CellReference(cellRef);
            if (!string.IsNullOrEmpty(cr.SheetName))
            {
                var newSheet = sheet.Workbook.GetSheet(cr.SheetName);
                return newSheet.GetCell(cr.Row, cr.Col);
            }

            return sheet.GetCell(cr.Row, cr.Col);
        }
        
        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <returns></returns>
        public static ICell CreateCell(this ISheet sheet, string cellRef)
        {
            var cr = new CellReference(cellRef);
            if (!string.IsNullOrEmpty(cr.SheetName))
            {
                var newSheet = sheet.Workbook.GetSheet(cr.SheetName);
                return newSheet.CreateCell(cr.Row, cr.Col);
            }

            return sheet.CreateCell(cr.Row, cr.Col);
        }

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static ICell CreateCell(this ISheet sheet, int row, int col)
        {
            var sheetRow = sheet.GetRow(row) ?? sheet.CreateRow(row);
            return sheetRow.CreateCell(col);
        }

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="firstRow"></param>
        /// <param name="lastRow"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <returns></returns>
        public static ICell CreateCell(this ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            var cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            sheet.AddMergedRegion(cellRangeAddress);
            return sheet.GetCell(firstRow, firstCol);
        }

        /// <summary>
        /// 获取或创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static ICell GetOrCreateCell(this ISheet sheet, int row, int cell)
        {
            return sheet.GetCell(row, cell) ?? sheet.CreateCell(row, cell);
        }

        /// <summary>
        /// 获取或创建单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <returns></returns>
        public static ICell GetOrCreateCell(this ISheet sheet, string cellRef)
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
        public static ICell SetCell(this ISheet sheet, string cellRef, bool data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.SetCellValue(data);
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static ICell SetCell(this ISheet sheet, string cellRef, double data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.SetCellValue(data);
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellRef"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static ICell SetCell(this ISheet sheet, string cellRef, string data)
        {
            var cell = sheet.GetOrCreateCell(cellRef);
            cell.SetCellValue(data);
            return cell;
        }

        /// <summary>
        /// 获取Sheet的总行数
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static int GetRowCount(this ISheet sheet)
        {
            return sheet.PhysicalNumberOfRows;
        }

        #endregion

        #region Row Extensions

        /// <summary>
        /// 获取指定标题<paramref name="text"/>的列 Index
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="text">要匹配的列内容</param>
        /// <param name="defaultIndex">如果没有匹配默认返回的Index</param>
        /// <param name="otherTexts">除过<paramref name="text"/>的其它匹配内容</param>
        /// <returns></returns>
        public static int GetIndex(this IRow row, string text, int defaultIndex = -1, params string[] otherTexts)
        {
            for (int i = 0; i < row.LastCellNum; i++)
            {
                var cellValue = row.GetCell(i)?.StringCellValue;
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
        /// 读取单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="calculate"></param>
        /// <returns></returns>
        public static object GetData(this ICell cell, bool calculate = true)
        {
            if (cell == null)
            {
                return null;
            }
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    {
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            return cell.DateCellValue;
                        }
                        return cell.NumericCellValue;
                    }
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Formula:
                    if (!calculate)
                    {
                        return cell.CellFormula;
                    }
                    try
                    {
                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.Blank:
                                return null;
                            case CellType.Boolean:
                                return cell.BooleanCellValue;
                            case CellType.Numeric:
                                {
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        return cell.DateCellValue;
                                    }
                                    return cell.NumericCellValue;
                                }
                            case CellType.String:
                                return cell.StringCellValue;
                            case CellType.Error:
                                return cell.ErrorCellValue;
                        }
                        return null;
                    }
                    catch
                    {
                        return cell.StringCellValue;
                    }
                default:
                    return cell.CellFormula;
            }
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static ICell SetValue(this ICell cell, string data)
        {
            cell.SetCellValue(data);
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static ICell SetValue(this ICell cell, double data)
        {
            cell.SetCellValue(data);
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static ICell SetValue(this ICell cell, bool data)
        {
            cell.SetCellValue(data);
            return cell;
        }

        /// <summary>
        /// 设置单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static ICell SetValue(this ICell cell, DateTime data)
        {
            cell.SetCellValue(data);
            return cell;
        }

        /// <summary>
        /// 设置单元格备注
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="comment"></param>
        /// <returns></returns>
        public static ICell SetComment(this ICell cell, string comment)
        {
            var patr = cell.Sheet.CreateDrawingPatriarch();
            if (patr is XSSFDrawing)
            {
                cell.CellComment = patr.CreateCellComment(new XSSFClientAnchor(0, 0, 0, 0, cell.Address.Column + 1, cell.Address.Row + 1, cell.Address.Column + 4, cell.Address.Row + 4));
                cell.CellComment.String = new XSSFRichTextString(comment);
                return cell;
            }
            if (patr is HSSFPatriarch)
            {
                cell.CellComment = patr.CreateCellComment(new HSSFClientAnchor(0, 0, 0, 0, cell.Address.Column + 1, cell.Address.Row + 1, cell.Address.Column + 4, cell.Address.Row + 4));
                cell.CellComment.String = new HSSFRichTextString(comment);
                return cell;
            }

            return cell;
        }

        /// <summary>
        /// 设置单元格边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="borderStype"></param>
        /// <returns></returns>
        public static ICell SetBorder(this ICell cell, BorderStyle borderStype = BorderStyle.Thin)
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
        public static ICell SetBorder(this ICell cell, BorderStyle borderTop, BorderStyle borderRight, BorderStyle borderBottom, BorderStyle borderLeft)
        {
            var cellStyle = cell.Sheet.Workbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(cell.CellStyle);
            cellStyle.BorderTop = borderTop;
            cellStyle.BorderRight = borderRight;
            cellStyle.BorderBottom = borderBottom;
            cellStyle.BorderLeft = borderLeft;
            cell.CellStyle = cellStyle;
            return cell;
        }

        /// <summary>
        /// 设置单元格边框
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="fontAction"></param>
        /// <returns></returns>
        public static ICell SetFont(this ICell cell, Action<IFont> fontAction)
        {
            var font = cell.Sheet.Workbook.CreateFont();
            fontAction.Invoke(font);

            var cellStyle = cell.Sheet.Workbook.CreateCellStyle();
            cellStyle.CloneStyleFrom(cell.CellStyle);
            cellStyle.SetFont(font);

            cell.CellStyle = cellStyle;

            return cell;
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public static ICell SetCellStyle(this ICell cell, ICellStyle cellStyle)
        {
            cell.CellStyle = cellStyle;
            return cell;
        }

        #endregion

    }
}
