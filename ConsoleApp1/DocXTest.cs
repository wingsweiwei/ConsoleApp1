using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace ConsoleApp1;

internal class DocXTest
{
    public void Run()
    {
        string path = @"C:\Test\SlopeInspectionReporter\template.docx";
        string outputPath = @"C:\Test\SlopeInspectionReporter\output.docx";
        var doc = DocX.Load(path);
        foreach (var table in doc.Tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var paragraph in cell.Paragraphs)
                    {
                        if (paragraph.Text.Contains("General View"))
                        {
                            Console.WriteLine(paragraph.Text);
                        }
                    }
                }
            }
        }
        foreach (var paragraph in doc.Paragraphs.Where(p => p.Text.Contains("General View")))
        {
            Console.WriteLine(paragraph.Text);
        }

        doc.SaveAs(outputPath);
    }
}
