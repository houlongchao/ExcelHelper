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
            _excelHelper.ExportSheet("test", datas)
                .ExportSheet("test2", data2)
                .ExportSheet("test3", data3, new ExcelHelper.Settings.ExportSetting()
                {
                    IgnoreProperties = new List<string>() { nameof(DemoIO.A), nameof(DemoIO.B) }
                })
                .SetSheetIndex("test3", 1);
            var bytes = _excelHelper.ToBytes();
            File.WriteAllBytes("test.xlsx", bytes);
        }
    }
}
