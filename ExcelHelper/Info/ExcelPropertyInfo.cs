using System.Collections.Generic;
using System.Reflection;

namespace ExcelHelper
{
    /// <summary>
    /// Excel 属性信息
    /// </summary>
    public class ExcelPropertyInfo
    {
        /// <summary>
        /// 字段属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; private set; }

        #region 导入

        /// <summary>
        /// 导入头
        /// </summary>
        public IEnumerable<ImportHeaderAttribute> ImportHeaders { get; private set; }

        /// <summary>
        /// 导入映射
        /// </summary>
        public IEnumerable<ImportMapperAttribute> ImportMappers { get; private set; }

        /// <summary>
        /// 导入限制
        /// </summary>
        public ImportLimitAttribute ImportLimit { get; set; }

        #endregion

        #region 导出

        /// <summary>
        /// 导出头
        /// </summary>
        public ExportHeaderAttribute ExportHeader { get; private set; }

        /// <summary>
        /// 导出映射
        /// </summary>
        public IEnumerable<ExportMapperAttribute> ExportMappers { get; private set; }

        /// <summary>
        /// 忽略导出，如果为null则导出，不为null则不导出
        /// </summary>
        public ExportIgnoreAttribute ExportIgnore { get; set; }

        #endregion

        /// <summary>
        /// Excel 属性信息
        /// </summary>
        /// <param name="propertyInfo"></param>
        public ExcelPropertyInfo(PropertyInfo propertyInfo)
        {
            PropertyInfo = propertyInfo;
        }

        /// <summary>
        /// Excel 属性信息
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <param name="importHeaders"></param>
        /// <param name="importMappers"></param>
        public ExcelPropertyInfo(PropertyInfo propertyInfo, IEnumerable<ImportHeaderAttribute> importHeaders, IEnumerable<ImportMapperAttribute> importMappers)
        {
            PropertyInfo = propertyInfo;
            ImportHeaders = importHeaders;
            ImportMappers = importMappers;
        }

        /// <summary>
        /// Excel 属性信息
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <param name="exportHeader"></param>
        /// <param name="exportMappers"></param>
        public ExcelPropertyInfo(PropertyInfo propertyInfo, ExportHeaderAttribute exportHeader, IEnumerable<ExportMapperAttribute> exportMappers)
        {
            PropertyInfo = propertyInfo;
            ExportHeader = exportHeader;
            ExportMappers = exportMappers;
        }
    }
}
