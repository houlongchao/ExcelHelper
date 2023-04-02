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
                Image = File.ReadAllBytes("D:\\050.jpg"),
            });
            var data2 = new List<DemoIO>(datas);
            data2.Add(new DemoIO()
            {
                A = "data2",
                B = "",
                C = "c",
                Status = Status.B,
                ImageName = "050.jpg",
                Image = File.ReadAllBytes("D:\\050.jpg"),
            });
            var data3 = new List<DemoIO>(data2);
            data3.Add(new DemoIO()
            {
                A = "data3",
                B = "",
                C = "c",
                Status = Status.B,
                ImageName = "050.jpg",
                Image = File.ReadAllBytes("D:\\050.jpg"),
            });

            _excelHelper.CreateExcelSheet("aaa").AppendData(data2).AppendEmptyRow().AppendData(data2).AppendData(data2, false);

            var setting01 = new ExportSetting();
            setting01.AddIgnoreProperties(nameof(DemoIO.A), nameof(DemoIO.B));
            var setting02 = new ExportSetting();
            setting02.AddIgnoreProperties(nameof(DemoIO.A), nameof(DemoIO.B));
            setting02.AddIncludeProperties(nameof(DemoIO.Date), nameof(DemoIO.B));
            setting02.AddTitleMapping(nameof(DemoIO.Date), "日期");
            setting02.AddTitleComment(nameof(DemoIO.Date), "日期备注");

            _excelHelper
                .ExportSheet("test", datas)
                .ExportSheet("test2", data2, setting01)
                .ExportSheet("test3", data3, setting02)
                .SetSheetIndex("test3", 1);
            var bytes = _excelHelper.ToBytes();
            File.WriteAllBytes("test.xlsx", bytes);
        }
    }
}
