using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1;

internal class ClosedXMLTest
{
    public void Test()
    {
        string folder = @"C:\Test\AIS";
        using var newWorkbook = new XLWorkbook();
        foreach (var filePath in Directory.GetFiles(folder, "*.xlsx"))
        {
            using var workbook = new XLWorkbook(filePath);
            foreach (var sheet in workbook.Worksheets)
            {
                sheet.CopyTo(newWorkbook, sheet.Name);
            }
        }
        newWorkbook.SaveAs(@"C:\Test\AIS\Merged.xlsx");
    }
}
