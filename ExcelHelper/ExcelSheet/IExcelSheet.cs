using ExcelHelper.Settings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        /// 获取数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        List<T> GetData<T>() where T : new();

        /// <summary>
        /// 获取总行数
        /// </summary>
        /// <returns></returns>
        int GetRowCount();
    }
}
