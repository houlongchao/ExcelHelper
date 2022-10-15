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
        /// 保存为字节数据
        /// </summary>
        /// <returns></returns>
        byte[] ToBytes();
    }
}
