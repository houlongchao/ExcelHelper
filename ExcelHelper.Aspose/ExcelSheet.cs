using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExcelHelper.Aspose
{
    /// <summary>
    /// Excel Sheet
    /// </summary>
    public class ExcelSheet : BaseExcelSheet
    {
        private readonly Worksheet _sheet;

        /// <summary>
        /// Aspose Worksheet
        /// </summary>
        public Worksheet Sheet => _sheet;

        /// <summary>
        /// Excel Sheet
        /// </summary>
        /// <param name="sheet"></param>
        public ExcelSheet(Worksheet sheet)
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
            return _sheet.GetRow(rowIndex).LastDataCell.Column + 1;
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
            var cell = _sheet.GetOrCreateCell(rowIndex, colIndex).SetValue(value);
            if (value is DateTime dt && DateTime.MinValue != dt)
            {
                SetFormat(cell, "yyyy/MM/dd HH:mm:ss", true);
            }
        }

        /// <inheritdoc/>
        public override void SetValue(string cellAddress, object value)
        {
            var cell = _sheet.GetOrCreateCell(cellAddress).SetValue(value);
            if (value is DateTime dt && DateTime.MinValue != dt)
            {
                SetFormat(cell, "yyyy/MM/dd HH:mm:ss", true);
            }
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

        private readonly IDictionary<string, Style> _styles = new Dictionary<string, Style>();
        
        /// <inheritdoc/>
        public override void SetFormat(int rowIndex, int colIndex, string format, bool cacheFormat = false)
        {
            var cell = _sheet.GetOrCreateCell(rowIndex, colIndex);
            SetFormat(cell, format, cacheFormat);
        }

        private void SetFormat(Cell cell, string format, bool cacheFormat = false)
        {
            if (cacheFormat)
            {
                if (_styles.TryGetValue(format, out var style))
                {
                    cell.SetCellStyle(style);
                }
                else
                {
                    var cellStyle = cell.Worksheet.Workbook.CreateStyle();
                    cellStyle.Copy(cell.GetStyle());
                    cellStyle.Custom = format;

                    _styles[format] = cellStyle;
                    cell.SetCellStyle(cellStyle);
                }
            }
            else
            {
                cell.SetDataFormat(format);
            }
        }

        /// <inheritdoc/>
        public override void SetAutoSizeColumn(int colIndex)
        {
            _sheet.AutoFitColumn(colIndex);
        }

        /// <inheritdoc/>
        public override void SetColumnWidth(int colIndex, int width)
        {
            _sheet.Cells.SetColumnWidth(colIndex, width);
        }

        /// <inheritdoc/>
        public override void SetFont(int rowIndex, int colIndex, string colorName = "Black", int fontSize = 12, bool isBold = true)
        {
            _sheet.GetOrCreateCell(rowIndex, colIndex).SetFont(font =>
                    {
                        font.Size = fontSize;
                        font.IsBold = isBold;
                        font.Color = Color.FromName(colorName);
                    });
        }

        /// <inheritdoc/>
        public override void SetValidationData(int firstRowIndex, int lastRowIndex, int firstColIndex, int lastColIndex, string[] explicitListValues)
        {
            var area = new CellArea();
            area.StartRow = firstRowIndex;
            area.EndRow = lastRowIndex;
            area.StartColumn = firstColIndex;
            area.EndColumn = lastColIndex;

            var index = _sheet.Validations.Add(area);
            var validation = _sheet.Validations[index];
            validation.Type = ValidationType.List;
            validation.Operator = OperatorType.Between;
            validation.Formula1 = string.Join(", ", explicitListValues);
            validation.ShowError = true;
            validation.AlertStyle = ValidationAlertType.Stop;
            validation.ErrorTitle = "请从指定列表中选择值";
            validation.ErrorMessage = string.Join(", ", explicitListValues);
        }
    }
}
