using System;
using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// 导入映射扩展方法
    /// </summary>
    public static class ImportMapperExtensions
    {
        /// <summary>
        /// 将显示值映射为实际值
        /// </summary>
        /// <param name="mappers">映射器</param>
        /// <param name="display">显示值</param>
        /// <returns></returns>
        public static object MappedToActual(this IEnumerable<ImportMapperAttribute> mappers, object display)
        {
            if (mappers == null)
            {
                return display;
            }

            foreach (var mapper in mappers)
            {
                if (display is DateTime dt && mapper.Display == dt.ToString("yyyy-MM-dd HH:mm:ss"))
                {
                    return mapper.Actual;
                }
                else if(display is Boolean b && mapper.Display == b.ToString().ToUpper())
                {
                    return mapper.Actual;
                }
                else if (mapper.Display == display.ToString())
                {
                    return mapper.Actual;
                }
            }

            return display;
        }
    }
}
