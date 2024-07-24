using System;

namespace ExcelHelper
{
    /// <summary>
    /// Excel 帮助类基类
    /// </summary>
    public abstract class BaseExcelTempHelper : IExcelTempHelper
    {
        /// <summary>
        /// 释放资源
        /// </summary>
        public virtual void Dispose() { }

        /// <inheritdoc/>
        public T GetData<T>(string filePath, string sheetName = null, TempSetting tempSetting = null) where T : new()
        {
            return GetExcelSheet(filePath, sheetName).GetTempData<T>(tempSetting);
        }

        /// <inheritdoc/>
        public byte[] SetData<T>(string tempPath, T data, string sheetName = null, TempSetting tempSetting = null) where T : new()
        {
            var excelSheet =  GetExcelSheet(tempPath, sheetName).SetTempData(data, tempSetting);
            return excelSheet.ToBytes();
        }

        /// <summary>
        /// 获取指定Excel文件指定Sheet页信息
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public abstract IExcelSheet GetExcelSheet(string filePath, string sheetName);
    }
}
