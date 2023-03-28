using ExcelHelper;
using System;

namespace ExcelHelperTest
{
    /// <summary>
    /// 导入导出测试模型
    /// </summary>
    [ImportUnique(nameof(A), nameof(B))]
    public class DemoIO
    {
        [ImportHeader("A", IsRequired = true, IsUnique = false)]
        [ImportHeader("AA", Trim = Trim.Start)]
        [ExportHeader("A2", ColorName = "Red")]
        public string A { get; set; }

        [ImportHeader("B")]
        [ImportHeader("BB")]
        [ExportHeader("B2")]
        public string B { get; set; }

        [ImportHeader("C")]
        [ImportHeader("CC")]
        [ImportMapper("A3", "b")]
        [ImportLimit("A3", true, 123)]
        [ExportHeader("C2", Comment = "备注")]
        [ExportMapper("a", "Aa")]
        [ExportMapper("b", "Ab")]
        [ExportMapper("c", "Ac")]
        public string C { get; set; }

        [ExportHeader("日期", ColumnWidth = 30)]
        public DateTime DateTime { get; set; }

        [ExportHeader("日期2", ColumnWidth = 30, Format = "yyyy/MM/dd")]
        public DateTime? DateTime2 { get; set; }

        [ExportIgnore]
        public DateTime Date { get; set; }

        [ImportLimit(-0.123)]
        [ExportHeader("数字", Format = "0.0")]
        public double Number { get; set; }

        public bool Boolean { get; set; }

        public string Formula { get; set; }

        [ExportMapper(ExcelHelperTest.Status.A, "AA")]
        [ExportMapper(null, "")]
        [ExportMapperElse("else")]
        public Status? Status { get; set; }

        public string ImageName { get; set; }

        [ExportHeader("图片", IsImage = true)]
        [ImportHeader("图片", IsImage = true)]
        public byte[] Image { get; set; }

    }

    public enum Status
    {
        A = 0,
        B = 1,
    }


    /// <summary>
    /// 参数值导入导出模型
    /// </summary>
    public class ParamItemIO
    {
        /// <summary>
        /// 编码
        /// </summary>
        [ExportHeader("编码")]
        [ImportHeader("编码", IsRequired = true, IsUnique = true)]
        public string Code { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        [ExportHeader("中文名称")]
        [ImportHeader("中文名称", IsRequired = true)]
        public string Name { get; set; }

        /// <summary>
        /// 英文名称
        /// </summary>
        [ExportHeader("英文名称")]
        [ImportHeader("英文名称")]
        public string NameEn { get; set; }

        /// <summary>
        /// 描述说明
        /// </summary>
        [ExportHeader("描述说明")]
        [ImportHeader("描述说明")]
        public string Desc { get; set; }
    }
}
