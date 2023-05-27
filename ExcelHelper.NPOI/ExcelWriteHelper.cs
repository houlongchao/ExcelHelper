using NPOI.SS.UserModel;
using System;

namespace ExcelHelper.NPOI
{
    /// <summary>
    /// Excel 写入帮助类
    /// </summary>
    public class ExcelWriteHelper : BaseExcelWriteHelper
    {
        private readonly IWorkbook _excel;

        /// <summary>
        /// NPOI IWorkbook
        /// </summary>
        public IWorkbook Excel => _excel;

        /// <summary>
        /// Excel 写入帮助类
        /// </summary>
        public ExcelWriteHelper()
        {
            _excel = NpoiHelper.CreateExcel();
        }

        /// <inheritdoc/>
        public override void Dispose()
        {
            _excel.Close();
            base.Dispose();
        }

        /// <summary>
        /// 创建一个 Excel Sheet 页
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
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

            _excel.SetSheetOrder(sheetName, index);

            return this;
        }

        /// <inheritdoc/>
        public override byte[] ToBytes()
        {
            return _excel.ToBytes();
        }
    }
}
