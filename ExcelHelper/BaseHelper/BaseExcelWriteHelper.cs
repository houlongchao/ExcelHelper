using System;
using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// Excel 帮助类基类
    /// </summary>
    public abstract class BaseExcelWriteHelper : IExcelWriteHelper
    {
        /// <summary>
        /// 释放资源
        /// </summary>
        public virtual void Dispose() { }

        /// <summary>
        /// 导出 Sheet 数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="datas"></param>
        /// <returns></returns>
        public IExcelWriteHelper ExportSheet<T>(string sheetName, IEnumerable<T> datas) where T : new()
        {
            CreateExcelSheet(sheetName).AppendData(datas);
            return this;
        }

        /// <summary>
        /// 导出 Sheet 数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="datas"></param>
        /// <param name="exportSetting"></param>
        /// <returns></returns>
        public IExcelWriteHelper ExportSheet<T>(string sheetName, IEnumerable<T> datas, ExportSetting exportSetting) where T : new()
        {
            CreateExcelSheet(sheetName).AppendData(datas, exportSetting);
            return this;
        }

        /// <inheritdoc/>
        public abstract IExcelWriteHelper SetSheetIndex(string sheetName, int index);

        /// <summary>
        /// 转为 byte 数据
        /// </summary>
        /// <returns></returns>
        public abstract byte[] ToBytes();


        /// <inheritdoc/>
        public abstract IExcelSheet CreateExcelSheet(string sheetName);

    }
}
