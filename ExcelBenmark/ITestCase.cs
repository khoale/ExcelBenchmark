namespace ExcelBenmark
{
    using System.Collections.Generic;

    public interface ITestCase
    {
        void ExportExcel(IList<dynamic> items, string fileName);
    }
}