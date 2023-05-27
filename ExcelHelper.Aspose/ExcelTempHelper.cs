namespace ExcelHelper.Aspose
{
    /// <summary>
    /// 模板Excel操作类
    /// </summary>
    public class ExcelTempHelper : BaseExcelTempHelper
    {
        /// <inheritdoc/>
        public override IExcelSheet GetExcelSheet(string filePath, string sheetName)
        {
            var excel = AsposeCellHelper.ReadExcel(filePath);
            if (excel == null)
            {
                return null;
            }
            var sheet = string.IsNullOrEmpty(sheetName) ?
                excel.GetSheetAt(0) :
                excel.GetSheet(sheetName);

            if (sheet == null)
            {
                return null;
            }

            return new ExcelSheet(sheet);
        }
    }
}
