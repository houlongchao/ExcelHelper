namespace ExcelHelper
{
    /// <summary>
    /// Excel 模板操作类
    /// </summary>
    public interface IExcelTempHelper
    {
        /// <summary>
        /// 获取模板数据
        /// </summary>
        /// <typeparam name="T">输出模型</typeparam>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetName">sheet页名称，默认第一个sheet页</param>
        /// <param name="tempSetting">模板设置</param>
        /// <returns></returns>
        T GetData<T>(string filePath, string sheetName = null, TempSetting tempSetting = null) where T: new();

        /// <summary>
        /// 设置模板数据
        /// </summary>
        /// <typeparam name="T">数据对象模型</typeparam>
        /// <param name="tempPath">模板文件路径</param>
        /// <param name="data">数据</param>
        /// <param name="sheetName">sheet页名称，默认第一个sheet页</param>
        /// <param name="tempSetting">模板设置</param>
        /// <returns></returns>
        byte[] SetData<T>(string tempPath, T data, string sheetName = null, TempSetting tempSetting = null) where T : new();
    }
}
