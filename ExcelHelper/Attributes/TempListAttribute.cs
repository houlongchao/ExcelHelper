using System;

namespace ExcelHelper
{
    /// <summary>
    /// 模板列表数据设置
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class TempListAttribute : Attribute
    {
        /// <summary>
        /// 数据类型
        /// </summary>
        public TempListType Type { get; set; }

        /// <summary>
        /// 开始位置
        /// </summary>
        public int StartIndex { get; set; }

        /// <summary>
        /// 结束位置
        /// </summary>
        public int EndIndex { get; set; }

        /// <summary>
        /// 模板列表数据设置
        /// </summary>
        /// <param name="type">行列表/列列表</param>
        /// <param name="startIndex">数据开始坐标（包含）</param>
        /// <param name="endIndex">数据结束坐标（包含）</param>
        public TempListAttribute(TempListType type, int startIndex, int endIndex)
        {
            Type = type;
            StartIndex = startIndex;
            EndIndex = endIndex;
        }
    }



    /// <summary>
    /// 列表类型
    /// </summary>
    public enum TempListType
    {
        /// <summary>
        /// 行
        /// </summary>
        Row,

        /// <summary>
        /// 列
        /// </summary>
        Column,
    }
}
