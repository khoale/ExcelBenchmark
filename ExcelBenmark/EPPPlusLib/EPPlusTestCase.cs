namespace ExcelBenmark.EPPPlusLib
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using OfficeOpenXml;

    public class EPPlusTestCase : ITestCase
    {
        public void ExportExcel(IList<dynamic> items, string fileName)
        {
            using (var excelPackage = new ExcelPackage(new FileStream(fileName, FileMode.CreateNew)))
            {
                var currentIndex = 0;
                var sheetIndex = 1;
                do
                {
                    var take = Math.Min(1000000, items.Count - currentIndex);
                    var currentSet = items.Skip(currentIndex).Take(take).ToList();
                    currentIndex += currentSet.Count;

                    var sheet = excelPackage.Workbook.Worksheets.Add("Sheet " + sheetIndex);
                    var arrayData = new List<object[]>();
                    foreach (var item in currentSet)
                    {
                        arrayData.Add(((IDictionary<string, object>)item).Values.ToArray());
                    }

                    sheet.Cells["A1"].LoadFromArrays(arrayData);

                    sheetIndex++;
                    excelPackage.Save();
                }
                while (currentIndex < items.Count);
            }
        }
    }
}