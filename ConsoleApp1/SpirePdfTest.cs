using Spire.Pdf.Utilities;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class SpirePdfTest
    {
        public void Run()
        {
            // https://www.slope.landsd.gov.hk/smris/map#
            string folder = @"C:\Work\SlopeInspectionWebApi\SlopeInspectionWebApi\Resources";
            foreach (var file in Directory.GetFiles(folder, "*.pdf"))
            {
                string fileFolder = Path.Combine(folder, Path.GetFileNameWithoutExtension(file));
                Directory.CreateDirectory(fileFolder);
                // 创建PdfDocument实例
                using PdfDocument pdf = new PdfDocument();

                // 载入PDF文档
                pdf.LoadFromFile(file);

                // 创建PdfImageHelper对象
                PdfImageHelper imageHelper = new PdfImageHelper();

                // 遍历PDF文档中的页面
                for (int i = 0; i < pdf.Pages.Count; i++)
                {
                    // 获取当前页面
                    PdfPageBase page = pdf.Pages[i];
                    // 获取当前页面的图片信息
                    PdfImageInfo[] imageInfos = imageHelper.GetImagesInfo(page);
                    // 遍历页面中的图片
                    for (int j = 0; j < imageInfos.Length; j++)
                    {
                        // 保存当前图片
                        string fileName = $"Page{i + 1}-Image{j + 1}.png";
                        using var fileStream = new FileStream(Path.Combine(fileFolder, fileName), FileMode.Create);
                        imageInfos[j].Image.CopyTo(fileStream);
                    }
                }
            }
        }
    }
}
