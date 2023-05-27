using System.Collections.Generic;

namespace ExcelHelper
{
    /// <summary>
    /// 模板配置
    /// </summary>
    public class TempSetting : BaseImportSetting
    {
        /// <summary>
        /// 属性单元格位置 (<c>nameof(A)</c>, <c>cellAddress</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>cellAddress</c> : Excel中单元格位置</para>
        /// </summary>
        public Dictionary<string, string> CellAddress { get; private set; } = new Dictionary<string, string>();

        /// <summary>
        /// 列表属性 (<c>nameof(A)</c>, <see cref="TempListSetting"/>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><see cref="TempListSetting"/> : 模板列表设置</para>
        /// </summary>
        public Dictionary<string, TempListSetting> ListSettings { get; private set; } = new Dictionary<string, TempListSetting>();

        #region Add

        /// <summary>
        /// 添加属性单元格位置
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="cellAddress">Excel中单元格位置</param>
        public void AddCellAddress(string propertyName, string cellAddress)
        {
            CellAddress[propertyName] = cellAddress;
        }

        /// <summary>
        /// 添加列表属性设置
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="type">行列表/列列表</param>
        /// <param name="startIndex">数据开始坐标（包含）</param>
        /// <param name="endIndex">数据结束坐标（包含）</param>
        public TempListSetting AddTempListSetting(string propertyName, TempListType type, int startIndex, int endIndex)
        {
            var tempListSetting = new TempListSetting()
            {
                Type = type,
                StartIndex = startIndex,
                EndIndex = endIndex,
            };
            ListSettings[propertyName] = tempListSetting;

            return tempListSetting;
        }

        #endregion
    }

    /// <summary>
    /// 模板列表设置
    /// </summary>
    public class TempListSetting : BaseImportSetting
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
        /// 列表数据位置 (<c>nameof(A)</c>, <c>itemIndex</c>)
        /// <para><c>nameof(A)</c> : 对象的指定属性A的名称</para>
        /// <para><c>itemIndex</c> : 列表数据位置</para>
        /// </summary>
        public Dictionary<string, int> ItemIndexs { get; private set; } = new Dictionary<string, int>();


        #region Add

        /// <summary>
        /// 添加列表数据位置
        /// </summary>
        /// <param name="propertyName">对象属性名称</param>
        /// <param name="itemIndex">列表数据位置</param>
        public void AddItemIndex(string propertyName, int itemIndex)
        {
            ItemIndexs[propertyName] = itemIndex;
        }

        #endregion
    }
}