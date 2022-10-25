using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// Excel 写入帮助类
    /// </summary>
    public interface IExcelWriteHelper
    {
        /// <summary>
        /// 导出数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        IExcelWriteHelper ExportSheet<T>(string sheetName, IEnumerable<T> datas) where T : new();

        /// <summary>
        /// 设置Sheet的位置
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        IExcelWriteHelper SetSheetIndex(string sheetName, int index);

        /// <summary>
        /// 保存为字节数据
        /// </summary>
        /// <returns></returns>
        byte[] ToBytes();
    }
}
