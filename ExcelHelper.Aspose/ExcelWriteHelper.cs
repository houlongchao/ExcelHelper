﻿using Aspose.Cells;

namespace ExcelHelper.Aspose
{
    /// <summary>
    /// Excel 写入帮助类
    /// </summary>
    public class ExcelWriteHelper : BaseExcelWriteHelper
    {
        private readonly Workbook _excel;

        /// <summary>
        /// Aspose Workbook
        /// </summary>
        public Workbook Excel => _excel;

        /// <summary>
        /// Excel 写入帮助类
        /// </summary>
        public ExcelWriteHelper() 
        {
            _excel = AsposeCellHelper.CreateExcel();
            _excel.Worksheets.Clear();
        }

        /// <inheritdoc/>
        public override void Dispose()
        {
            _excel.Dispose();
            base.Dispose();
        }

        /// <inheritdoc/>
        public override IExcelSheet CreateExcelSheet(string sheetName)
        {
            var sheet = _excel.CreateSheet(sheetName);
            return new ExcelSheet(sheet);
        }

        /// <inheritdoc/>
        public override IExcelWriteHelper SetSheetIndex(string sheetName, int index)
        {
            var sheet = _excel.GetSheet(sheetName);
            if (sheet == null)
            {
                return this;
            }

            sheet.MoveTo(index);

            return this;
        }

        /// <inheritdoc/>
        public override byte[] ToBytes()
        {
            return _excel.ToBytes();
        }
    }
}
