using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System.Collections;
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
        public override void SetValue(int rowIndex, int colIndex, object value)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetValue(value);
        }

        /// <inheritdoc/>
        public override void SetValue(string cellAddress, object value)
        {
            _sheet.GetOrCreateCell(cellAddress).SetValue(value);
        }

        /// <inheritdoc/>
        public override void SetImage(int rowIndex, int colIndex, byte[] value)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetImage(value);
        }

        /// <inheritdoc/>
        public override void SetImage(string cellAddress, byte[] value)
        {
            _sheet.GetOrCreateCell(cellAddress).SetImage(value);
        }

        /// <inheritdoc/>
        public override void SetComment(int rowIndex, int colIndex, string comment)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetComment(comment);
        }

        private readonly IDictionary<string, ICellStyle> _styles = new Dictionary<string, ICellStyle>();

        /// <inheritdoc/>
        public override void SetFormat(int rowIndex, int colIndex, string format, bool cacheFormat = false)
        {
            if (cacheFormat)
            {
                if (_styles.TryGetValue(format, out var style))
                {
                    _sheet.GetOrCreateCell(rowIndex, colIndex).SetCellStyle(style);
                }
                else
                {
                    var cell = _sheet.GetOrCreateCell(rowIndex, colIndex);
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
                _sheet.GetOrCreateCell(rowIndex, colIndex).SetDataFormat(format);
            }
        }

        /// <inheritdoc/>
        public override void SetAutoSizeColumn(int colIndex)
        {
            _sheet.AutoSizeColumn(colIndex);
        }

        /// <inheritdoc/>
        public override void SetColumnWidth(int colIndex, int width)
        {
            _sheet.SetColumnWidth(colIndex, width * 256);
        }

        /// <inheritdoc/>
        public override void SetFont(int rowIndex, int colIndex, string colorName = "Black", int fontSize = 12, bool isBold = true)
        {
            var indexedColor = IndexedColors.ValueOf(colorName);
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetFont(font =>
            {
                font.FontHeight = fontSize * 20;
                font.IsBold = isBold;
                font.Color = indexedColor?.Index ?? IndexedColors.Black.Index;
            });
        }

        /// <inheritdoc/>
        public override void SetValidationData(int firstRowIndex, int lastRowIndex, int firstColIndex, int lastColIndex, string[] explicitListValues)
        {
            var constraint =  new XSSFDataValidationConstraint(explicitListValues);
            var addressList = new CellRangeAddressList(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex);
            var validation = _sheet.GetDataValidationHelper().CreateValidation(constraint, addressList);
            validation.SuppressDropDownArrow = true;
            validation.ShowErrorBox = true;
            validation.CreateErrorBox("请从指定列表中选择值", string.Join(", ", explicitListValues));
            _sheet.AddValidationData(validation);
        }
    }
}
