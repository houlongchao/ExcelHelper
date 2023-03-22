using System;
using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// 导出映射扩展方法
    /// </summary>
    public static class ExportMapperExtensions
    {
        /// <summary>
        /// 将实际值映射为显示数据
        /// </summary>
        /// <param name="mappers">映射器</param>
        /// <param name="actual">实际值</param>
        /// <returns></returns>
        public static object MappedToDisplay(this IEnumerable<ExportMapperAttribute> mappers, object actual)
        {
            if (mappers == null)
            {
                return actual;
            }

            foreach (var mapper in mappers)
            {
                if (actual is DateTime dt && dt.Equals(mapper.Actual))
                {
                    return mapper.Display;
                }
                else if(actual is Boolean b && b.Equals(mapper.Actual))
                {
                    return mapper.Display;
                }
                else if (actual is double d && d.Equals(Convert.ToDouble(mapper.Actual)))
                {
                    return mapper.Display;
                }
                else if (actual is float df && df.Equals(Convert.ToDouble(mapper.Actual)))
                {
                    return mapper.Display;
                }
                else if (actual is decimal dc && dc.Equals(Convert.ToDecimal(mapper.Actual)))
                {
                    return mapper.Display;
                }
                else if (actual is int di && di.Equals(Convert.ToInt32(mapper.Actual)))
                {
                    return mapper.Display;
                }
                else if (actual == mapper.Actual)
                {
                    return mapper.Display;
                }
            }

            return actual;
        }
    }
}
