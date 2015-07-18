namespace ExcelBenmark.NPOILib
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;

    public class NPOITestCase : ITestCase
    {
        public void ExportExcel(IList<dynamic> items, string fileName)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("Sheet1");

            for (int rownum = 0; rownum < items.Count; rownum++)
            {
                IRow row = worksheet.CreateRow(rownum);
                var item = items[rownum];
                var keys = item.Keys.ToArray();
                for (int index = 0; index < keys.Length; index++)
                {
                    var key = keys[index];
                    ICell cell = row.CreateCell(index);
                    var value = item[key];
                    if (value != null)
                    {
                        cell.SetCellValue(value.ToString());
                    }
                }
            }

            using (FileStream sw = File.Create(fileName))
            {
                workbook.Write(sw);
                sw.Close();
            }
        }
    }
}