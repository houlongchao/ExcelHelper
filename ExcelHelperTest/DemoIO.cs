using ExcelHelper;
using System;

namespace ExcelHelperTest
{
    /// <summary>
    /// 导入导出测试模型
    /// </summary>
    public class DemoIO
    {
        [ImportHeader("A")]
        [ImportHeader("AA")]
        [ExportHeader("A2")]
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

        [ExportMapper(Status.A, "AA")]
        public Status Status { get; set; }

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
}
