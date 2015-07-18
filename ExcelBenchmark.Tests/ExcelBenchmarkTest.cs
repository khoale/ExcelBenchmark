namespace ExcelBenchmark.Tests
{
    using System.IO;
    using System.Linq;
    using System.Threading;

    using ExcelBenchmark.Tests.Extensions;

    using ExcelBenmark.Data;
    using ExcelBenmark.EPPPlusLib;

    using NUnit.Framework;

    public class ExcelBenchmarkTest
    {
        private readonly DataProvider dataProvider;

        public ExcelBenchmarkTest()
        {
            dataProvider = new DataProvider();

            if (Directory.Exists("out"))
            {
                Directory.Delete("out", true);
            }
            Thread.Sleep(100);

            Directory.CreateDirectory("out");
        }

        [Test]
        public void TestProduct()
        {
            var products = dataProvider.GetProducts().ToList().Duplicate(10000);
            var epplusTestCase = new EPPlusTestCase();
            epplusTestCase.ExportExcel(products, "out/epplustest.xlsx");
        }
    }
}
