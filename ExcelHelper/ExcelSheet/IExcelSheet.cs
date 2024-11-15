using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// Excel Sheet
    /// </summary>
    public interface IExcelSheet
    {
        /// <summary>
        /// 追加数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="datas"></param>
        /// <param name="addTitle"></param>
        /// <returns></returns>
        IExcelSheet AppendData<T>(IEnumerable<T> datas, bool addTitle = true) where T : new();

        /// <summary>
        /// 追加数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="datas"></param>
        /// <param name="exportSetting"></param>
        /// <returns></returns>
        IExcelSheet AppendData<T>(IEnumerable<T> datas, ExportSetting exportSetting) where T : new();

        /// <summary>
        /// 追加空行
        /// </summary>
        /// <returns></returns>
        IExcelSheet AppendEmptyRow();

        /// <summary>
        /// 设置指定位置的数据
        /// </summary>
        /// <returns></returns>
        IExcelSheet SetValue(int rowIndex, int colIndex, object value);

        /// <summary>
        /// 设置指定位置的数据
        /// </summary>
        /// <returns></returns>
        IExcelSheet SetValue(string cellAddress, object value);

        /// <summary>
        /// 设置指定位置的图片数据
        /// </summary>
        /// <returns></returns>
        IExcelSheet SetImage(int rowIndex, int colIndex, byte[] value);

        /// <summary>
        /// 设置指定位置的图片数据
        /// </summary>
        /// <returns></returns>
        IExcelSheet SetImage(string cellAddress, byte[] value);

        /// <summary>
        /// 设置指定位置的备注信息
        /// </summary>
        IExcelSheet SetComment(int rowIndex, int colIndex, string comment);

        /// <summary>
        /// 设置指定位置的格式化字符串
        /// </summary>
        IExcelSheet SetFormat(int rowIndex, int colIndex, string format, bool cacheFormat = false);

        /// <summary>
        /// 设置指定位置的字体
        /// </summary>
        IExcelSheet SetFont(int rowIndex, int colIndex, string colorName = "Black", int fontSize = 12, bool isBold = true);

        /// <summary>
        /// 设置验证数据
        /// </summary>
        /// <param name="firstRowIndex"></param>
        /// <param name="lastRowIndex"></param>
        /// <param name="firstColIndex"></param>
        /// <param name="lastColIndex"></param>
        /// <param name="explicitListValues"></param>
        IExcelSheet SetValidationData(int firstRowIndex, int lastRowIndex, int firstColIndex, int lastColIndex, string[] explicitListValues);


        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="firstRow">起始行</param>
        /// <param name="firstCol">起始列</param>
        /// <param name="totalRows">总行数</param>
        /// <param name="totalColumns">总列数</param>
        /// <returns></returns>
        IExcelSheet MergedRegion(int firstRow, int firstCol, int totalRows, int totalColumns);

        /// <summary>
        /// 获取数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        List<T> GetData<T>(ImportSetting importSetting = null) where T : new();

        /// <summary>
        /// 获取总行数
        /// </summary>
        /// <returns></returns>
        int GetRowCount();

        /// <summary>
        /// 获取指定行的列数
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        int GetColumnCount(int rowIndex = 0);

        /// <summary>
        /// 保存为字节数据
        /// </summary>
        /// <returns></returns>
        byte[] ToBytes();

        /// <summary>
        /// 获取模板数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        T GetTempData<T>(TempSetting tempSetting = null) where T : new();

        /// <summary>
        /// 设置模板数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        IExcelSheet SetTempData<T>(T data, TempSetting tempSetting = null) where T : new();
    }
}
