using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;

namespace z.Office.Microsoft.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            using (var ms = new z.Office.Microsoft.ExcelWriter(false))
            {
                ms.FileName = Path.Combine(Environment.CurrentDirectory, Guid.NewGuid().ToString() + ".xls");

                ms.AddSheet("test");

                var row = ms.AddRow("test");

                ms.GetOrCreateCell(row, 0, "Hello World");
                ms.GetOrCreateCell(row, 1, int.MaxValue);
                var cell = ms.GetOrCreateCell(row, 2, DateTime.Now);
                ms.AddCellComment(cell, "Hello from world");

                ms.Save();

            }
        }

        [TestMethod]
        public void TestUpload()
        {
            var mfile = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "StudentTemplate-3fcbe412e7ed47a9a07c31f818318caf.xlsx");

            var xls = new ExcelReader(mfile);

            var sheet = xls.SheetsNames.First();
            var data = xls.ReadSheet(sheet);

            Assert.IsNotNull(data);
            Assert.AreEqual(1, data.Count);
            Assert.AreEqual("20210001", data[0]["Student No"]);

        }

        [TestMethod]
        public void ReadExcel()
        {
            var mfile = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "StudentTemplate-a7716fbb9fe845a5acf787ee82b02c73.xlsx");

            var mpc = File.ReadAllBytes(mfile);

            var xls = new ExcelReader(new MemoryStream(mpc), true);

            var sheet = xls.SheetsNames.First();
            var data = xls.ReadSheet(sheet, 5);

            Assert.IsNotNull(data);
            Assert.AreEqual(1, data.Count);
            Assert.AreEqual("Rizal", data[0]["Last Name"]);
        }
    }
}
