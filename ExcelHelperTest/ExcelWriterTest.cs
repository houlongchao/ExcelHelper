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
                Number = 0.123,
                User = new User()
                {
                    Name = "user1",
                    Age = 20,   
                },
                Users = new List<User>()
                {
                    new User()
                    {
                        Name = "user2",
                        Age = 21,
                    },
                    new User()
                    {
                        Name = "user3",
                        Age = 22,
                    },
                },
                UserDict = new Dictionary<string, User>()
                {
                    {"user4", new User()
                    {
                        Name = "user4",
                        Age = 23,
                    } },
                    {"user5", new User()
                    {
                        Name = "user5",
                        Age = 24,
                    } },
                }

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
            var data3 = new List<DemoIO>();
            data3.Add(new DemoIO()
            {
                A = "data3",
                B = "",
                C = "c",
                Status = Status.B,
                ImageName = "050.jpg",
                Image = File.ReadAllBytes("image.jpg"),
                OtherPropries = new Dictionary<string, object>
                {
                    { "a", "aa" },
                    { "b", true },
                }
            });

            //_excelHelper.CreateExcelSheet("aaa").AppendData(data2).AppendEmptyRow().AppendData(data2).AppendData(data2, false);

            var setting2 = new ExportSetting();
            setting2.AddIgnoreProperties(nameof(DemoIO.A), nameof(DemoIO.B));

            var setting3 = new ExportSetting();
            setting3.AddIgnoreProperties(nameof(DemoIO.A), nameof(DemoIO.B));
            setting3.AddIncludeProperties(nameof(DemoIO.Date), nameof(DemoIO.B), "OtherPropries.a", "OtherPropries.b", "OtherPropries.c");
            setting3.AddTitleMapping(nameof(DemoIO.Date), "日期");
            setting3.AddTitleComment(nameof(DemoIO.Date), "日期备注");
            setting3.AddTitleMapping("OtherPropries.a", "扩展属性A");
            //_excelHelper.ExportSheet("test", data3, setting3);
            var data4 = new List<Dictionary<string, object>>();
            data4.Add(new Dictionary<string, object>()
            {
                {"a", "aa" },
                {"b", true },
                {"c", DateTime.Now },
                {"d", 1.1 },
            });

            var setting4 = new ExportSetting();
            setting4.AddIncludeProperties("a", "c", "b", "d");
            setting4.AddTitleMapping("a", "字符串");
            setting4.AddTitleMapping("c", "日期");
            setting4.AddTitleComment("c", "日期备注");

            _excelHelper
                .ExportSheet("test", datas)
                .ExportSheet("test2", data2, setting2)
                .ExportSheet("test3", data3, setting3)
                .ExportSheet("test4", data4, setting4)
                .SetSheetIndex("test3", 1);

            var locationSheet = _excelHelper.CreateExcelSheet("location1");
            locationSheet.SetValue(0, 0, "A").MergedRegion(0, 0, 1, 5);
            locationSheet.SetValue(0, 5, "B").MergedRegion(0, 5, 1, 5);
            locationSheet.AppendData(datas, new ExportSetting()
            {
                ExportLocation = ExcelHelper.Settings.ExportLocation.LastRow,
            });
            locationSheet.AppendData(datas, new ExportSetting()
            {
                ExportLocation = ExcelHelper.Settings.ExportLocation.LastRow,
            });
            var locationSheet2 = _excelHelper.CreateExcelSheet("location2");
            locationSheet2.AppendData(datas, new ExportSetting()
            {
                ExportLocation = ExcelHelper.Settings.ExportLocation.Custom,
                ExportRowIndex = 2,
                ExportColumnIndex = 2,
            });
            var bytes = _excelHelper.ToBytes();
            File.WriteAllBytes("test.xlsx", bytes);
        }

        [Test]
        public void Test_ExportDictionary()
        {
            var datas = new List<Dictionary<string, object>>()
            {
                new Dictionary<string, object>()
                {
                    {"a", "a1" },
                    {"b", "b1" }
                },
                new Dictionary<string, object>()
                {
                    {"a", "a2" },
                    {"b", "b2" }
                },
            };
            var setting = new ExportSetting();
            setting.AddIncludeProperties("a", "b");  // 导出字典类型时必须指定
            setting.AddTitleMapping("a", "A");
            setting.AddTitleMapping("b", "B");  

            _excelHelper.ExportSheet("dict", datas, setting);
            var bytes = _excelHelper.ToBytes();
            File.WriteAllBytes("test_export_dict.xlsx", bytes);
        }

    }
}
