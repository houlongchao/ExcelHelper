using ExcelHelper;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelHelperTest
{
    public abstract class ExcelWriterTest
    {
        protected IExcelWriteHelper _excelHelper;

        [Test]
        public void Test_ExportSheet()
        {
            var datas = new List<DemoIO>();
            datas.Add(new DemoIO()
            {
                A = "a1",
                B = "b2",
                C = "a",
                DateTime = DateTime.Now,
                DateTime2 = DateTime.Now,
                Number = 0.123

            });
            datas.Add(new DemoIO()
            {
                A = "a1",
                B = "b2",
                C = "b",
                Status = null,
                Number = -123.3456
            });
            datas.Add(new DemoIO()
            {
                A = "a1",
                B = "b2",
                C = "C3",
                Status = Status.A
            });; 
            datas.Add(new DemoIO()
            {
                A = "a1",
                B = "b2aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
                C = "c",
                Status = Status.B,
                ImageName = "050.jpg",
                Image = File.ReadAllBytes("image.jpg"),
            });
            var data2 = new List<DemoIO>(datas);
            data2.Add(new DemoIO()
            {
                A = "data2",
                B = "",
                C = "c",
                Status = Status.B,
                ImageName = "050.jpg",
                Image = File.ReadAllBytes("image.jpg"),
            });
            var data3 = new List<DemoIO>(data2);
            data3.Add(new DemoIO()
            {
                A = "data3",
                B = "",
                C = "c",
                Status = Status.B,
                ImageName = "050.jpg",
                Image = File.ReadAllBytes("image.jpg"),
            });

            _excelHelper.CreateExcelSheet("aaa").AppendData(data2).AppendEmptyRow().AppendData(data2).AppendData(data2, false);

            var setting2 = new ExportSetting();
            setting2.AddIgnoreProperties(nameof(DemoIO.A), nameof(DemoIO.B));

            var setting3 = new ExportSetting();
            setting3.AddIgnoreProperties(nameof(DemoIO.A), nameof(DemoIO.B));
            setting3.AddIncludeProperties(nameof(DemoIO.Date), nameof(DemoIO.B));
            setting3.AddTitleMapping(nameof(DemoIO.Date), "日期");
            setting3.AddTitleComment(nameof(DemoIO.Date), "日期备注");

            var data4 = new List<Dictionary<string, object>>();
            data4.Add(new Dictionary<string, object>()
            {
                {"a", "aa" },
                {"b", true },
                {"c", DateTime.Now },
                {"d", 1.1 },
            });

            var setting4 = new ExportSetting();
            setting4.AddIgnoreProperties("b");
            setting4.AddIncludeProperties("a", "b", "c", "d");
            setting4.AddTitleMapping("a", "字符串");
            setting4.AddTitleMapping("c", "日期");
            setting4.AddTitleComment("c", "日期备注");

            _excelHelper
                .ExportSheet("test", datas)
                .ExportSheet("test2", data2, setting2)
                .ExportSheet("test3", data3, setting3)
                .ExportSheet("test4", data4, setting4)
                .SetSheetIndex("test3", 1);
            var bytes = _excelHelper.ToBytes();
            File.WriteAllBytes("test.xlsx", bytes);
        }
    }
}
