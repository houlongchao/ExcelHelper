using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;

namespace ExcelHelper.NPOI
{
    /// <summary>
    /// Excel Sheet
    /// </summary>
    public class ExcelSheet : BaseExcelSheet
    {
        private readonly ISheet _sheet;

        /// <summary>
        /// NPOI ISheet
        /// </summary>
        public ISheet Sheet => _sheet;

        /// <summary>
        /// Excel Sheet
        /// </summary>
        /// <param name="sheet"></param>
        public ExcelSheet(ISheet sheet)
        {
            _sheet = sheet;
        }

        /// <inheritdoc/>
        public override int GetRowCount()
        {
            return _sheet.GetRowCount();
        }

        /// <inheritdoc/>
        public override int GetColumnCount(int rowIndex = 0)
        {
            return _sheet.GetRow(rowIndex).LastCellNum;
        }

        /// <inheritdoc/>
        public override byte[] ToBytes()
        {
            return _sheet.Workbook.ToBytes();
        }

        /// <inheritdoc/>
        public override byte[] GetImage(int rowIndex, int colIndex)
        {
            return _sheet.GetOrCreateCell(rowIndex, colIndex).GetImage();
        }

        /// <inheritdoc/>
        public override byte[] GetImage(string cellAddress)
        {
            return _sheet.GetOrCreateCell(cellAddress).GetImage();
        }

        /// <inheritdoc/>
        public override object GetValue(int rowIndex, int colIndex)
        {
            return _sheet.GetOrCreateCell(rowIndex, colIndex).GetData();
        }

        /// <inheritdoc/>
        public override object GetValue(string cellAddress)
        {
            return _sheet.GetOrCreateCell(cellAddress).GetData();
        }

        /// <inheritdoc/>
        public override IExcelSheet SetValue(int rowIndex, int colIndex, object value)
        {
            var cell = _sheet.GetOrCreateCell(rowIndex, colIndex).SetValue(value);
            if (value is DateTime dt && DateTime.MinValue != dt)
            {
                SetFormat(cell, "yyyy/MM/dd HH:mm:ss", true);
            }
            return this;
        }

        /// <inheritdoc/>
        public override IExcelSheet SetValue(string cellAddress, object value)
        {
            var cell = _sheet.GetOrCreateCell(cellAddress).SetValue(value);
            if (value is DateTime dt && DateTime.MinValue != dt)
            {
                SetFormat(cell, "yyyy/MM/dd HH:mm:ss", true);
            }
            return this;
        }

        /// <inheritdoc/>
        public override IExcelSheet SetImage(int rowIndex, int colIndex, byte[] value)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetImage(value);
            return this;
        }

        /// <inheritdoc/>
        public override IExcelSheet SetImage(string cellAddress, byte[] value)
        {
            _sheet.GetOrCreateCell(cellAddress).SetImage(value);
            return this;
        }

        /// <inheritdoc/>
        public override IExcelSheet SetComment(int rowIndex, int colIndex, string comment)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetComment(comment);
            return this;
        }

        private readonly IDictionary<string, ICellStyle> _styles = new Dictionary<string, ICellStyle>();

        /// <inheritdoc/>
        public override IExcelSheet SetFormat(int rowIndex, int colIndex, string format, bool cacheFormat = false)
        {
            var cell = _sheet.GetOrCreateCell(rowIndex, colIndex);
            SetFormat(cell, format, cacheFormat);
            return this;
        }

        private void SetFormat(ICell cell, string format, bool cacheFormat = false)
        {
            if (cacheFormat)
            {
                if (_styles.TryGetValue(format, out var style))
                {
                    cell.SetCellStyle(style);
                }
                else
                {
                    var cellStyle = cell.Sheet.Workbook.CreateCellStyle();
                    cellStyle.CloneStyleFrom(cell.CellStyle);
                    var df = cell.Sheet.Workbook.CreateDataFormat();
                    cellStyle.DataFormat = df.GetFormat(format);
                    _styles[format] = cellStyle;
                    cell.SetCellStyle(cellStyle);
                }
            }
            else
            {
                // 该方式每次都会创建一个style，数据量大时会报错
                cell.SetDataFormat(format);
            }
        }

        /// <inheritdoc/>
        public override IExcelSheet SetAutoSizeColumn(int colIndex)
        {
            _sheet.AutoSizeColumn(colIndex);
            return this;
        }

        /// <inheritdoc/>
        public override IExcelSheet SetColumnWidth(int colIndex, int width)
        {
            _sheet.SetColumnWidth(colIndex, width * 256);
            return this;
        }

        /// <inheritdoc/>
        public override IExcelSheet SetFont(int rowIndex, int colIndex, string colorName = "Black", int fontSize = 12, bool isBold = true)
        {
            var indexedColor = IndexedColors.ValueOf(colorName);
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetFont(font =>
            {
                font.FontHeight = fontSize * 20;
                font.IsBold = isBold;
                font.Color = indexedColor?.Index ?? IndexedColors.Black.Index;
            });
            return this;
        }

        /// <inheritdoc/>
        public override IExcelSheet SetValidationData(int firstRowIndex, int lastRowIndex, int firstColIndex, int lastColIndex, string[] explicitListValues)
        {
            var constraint =  new XSSFDataValidationConstraint(explicitListValues);
            var addressList = new CellRangeAddressList(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex);
            var validation = _sheet.GetDataValidationHelper().CreateValidation(constraint, addressList);
            validation.SuppressDropDownArrow = true;
            validation.ShowErrorBox = true;
            validation.CreateErrorBox("请从指定列表中选择值", string.Join(", ", explicitListValues));
            _sheet.AddValidationData(validation);
            return this;
        }

        /// <inheritdoc/>
        public override IExcelSheet MergedRegion(int firstRow, int firstCol, int totalRows, int totalColumns)
        {
            _sheet.AddMergedRegion(new CellRangeAddress(firstRow, firstRow + totalRows - 1, firstCol, firstCol + totalColumns - 1));
            return this;
        }
    }
}
