using ExcelHelper;
using System;

namespace ExcelHelperTest
{
    /// <summary>
    /// 导入导出测试模型
    /// </summary>
    [ImportUniques(nameof(A), nameof(B))]
    //[ImportUniques(nameof(A), nameof(B), Message = "数据必须唯一提示")]
    public class DemoIO
    {
        [ImportHeader("A")]
        [ImportHeader("AA")]
        [ImportRequired]
        //[ImportRequired(Message = "数据必填提示")]
        [ImportUnique]
        //[ImportUnique(Message = "数据唯一提示")]
        [ImportTrim(Trim.All)]
        [ImportLimit("A1", "A2", "A3")]
        //[ImportLimit("A1", "A2", "A3", Message = "数据限制提示")]
        [ExportHeader("A2", ColorName = "Red")]
        public string A { get; set; }

        [ImportHeader("B")]
        [ImportHeader("BB")]
        [ImportRequired(Message = "数据B必填")]
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

        [ExportHeader("日期2", ColumnWidth = 30)]
        [ExportFormat("yyyy/MM/dd")]
        public DateTime? DateTime2 { get; set; }

        [ExportIgnore]
        public DateTime Date { get; set; }

        [ExportHeader("数字")]
        [ExportFormat("0.0")]
        public double Number { get; set; }

        public bool Boolean { get; set; }

        public string Formula { get; set; }

        [ExportMapper(ExcelHelperTest.Status.A, "AA")]
        [ExportMapper(null, "")]
        [ExportMapperElse("else")]
        public Status? Status { get; set; }

        public string ImageName { get; set; }

        [ExportHeader("图片")]
        [ImportHeader("图片")]
        [Image]
        public byte[] Image { get; set; }
    }

    public enum Status
    {
        A = 0,
        B = 1,
    }
}
