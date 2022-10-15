namespace ExcelHelper
{
    /// <summary>
    /// Excel Sheet 页信息
    /// </summary>
    public class ExcelSheetInfo
    {
        /// <summary>
        /// Sheet 名称
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Sheet 位置
        /// </summary>
        public int Index { get; private set; }

        /// <summary>
        /// 是否隐藏
        /// </summary>
        public bool IsHidden { get; private set; }

        /// <summary>
        /// Sheet 页信息
        /// </summary>
        /// <param name="index"></param>
        /// <param name="name"></param>
        /// <param name="isHidden"></param>
        public ExcelSheetInfo(int index, string name, bool isHidden)
        {
            Name = name;
            Index = index;
            IsHidden = isHidden;
        }

    }
}
