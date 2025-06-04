using Polly;
using Spire.Pdf;
using Spire.Pdf.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class SpirePdfTest
    {
        public async Task Run()
        {
            var sw = Stopwatch.StartNew();
            // https://www.slope.landsd.gov.hk/smris/map#

            string[] urls =
            [
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/C176&sd=&extent=830918.7621922652,814492.1966565147,831067.3068077349,814622.8263434854&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/C176",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/C175&sd=&extent=830976.2392545324,814416.5060836071,831124.783870002,814547.1357705778&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/C175",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/C386&sd=&extent=831077.9606922652,814260.6931565141,831226.5053077348,814391.3228434848&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/C386",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/C387&sd=&extent=831095.074946908,814269.6036699166,831243.6195623777,814400.2333568874&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/C387",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/FR149&sd=&extent=830946.82891358,814427.2550407065,831095.3735290497,814557.8847276772&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/FR149",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/R216&sd=&extent=830917.7171547695,814492.1966565148,831066.2617702391,814622.8263434855&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/R216",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/CR388&sd=&extent=831143.3781922652,814195.8281565148,831291.9228077348,814326.4578434855&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/CR388",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/FR43&sd=&extent=831063.4638845304,814065.6713130291,831360.5531154696,814326.9306869708&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/FR43",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/FR164&sd=&extent=831285.0111922651,814144.5476565144,831433.5558077347,814275.1773434852&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/FR164",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-C/C502&sd=&extent=831430.4206952475,814234.1222990087,831578.9653107171,814364.7519859794&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-C/C502",

                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SE-B/CR569&sd=&extent=842658.9746922652,814961.6736565146,842807.5193077348,815092.3033434853&imageWidth=898&imageHeight=700&lang=en&keyword=11SE-B/CR569",

                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-A/F63&sd=&extent=838865.6256922652,809626.4436565151,839014.1703077349,809757.0733434858&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-A/F63",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-D/C865&sd=&extent=836182.6776922649,812062.8221565138,836331.2223077345,812193.4518434845&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-D/C865",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-D/F207&sd=&extent=836056.8476922653,812128.4451565144,836205.3923077349,812259.0748434851&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-D/F207",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-D/FR225&sd=&extent=835979.9627907911,812263.5535756098,836128.5074062607,812394.1832625805&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-D/FR225",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-D/F336&sd=&extent=836496.2051922652,812637.7321565149,836644.7498077349,812768.3618434856&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-D/F336",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-D/R414&sd=&extent=835961.8011922651,812291.641156514,836110.3458077348,812422.2708434847&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-D/R414",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-A/F70&sd=&extent=838884.0651922651,809615.2291565144,839032.6098077347,809745.8588434851&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-A/F70",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-A/C120&sd=&extent=838802.2536854537,809614.3334100895,838950.7983009233,809744.9630970602&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-A/C120",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-A/C265&sd=&extent=838842.7115656469,809606.8688565483,838991.2561811166,809737.4985435191&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-A/C265",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-A/F62&sd=&extent=838629.1328845306,809578.5303130293,838926.2221154699,809839.789686971&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-A/F62",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-A/R115&sd=&extent=838443.6666922654,809822.7726565141,838592.2113077351,809953.4023434848&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-A/R115",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-A/R134&sd=&extent=838798.2016922651,809115.812156514,838946.7463077347,809246.4418434848&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-A/R134",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-D/CR1056&sd=&extent=835847.854692265,812293.3211565155,835996.3993077347,812423.9508434862&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-D/CR1056",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=11SW-D/CR1055&sd=&extent=835841.8830494321,812316.7598546349,835990.4276649017,812447.3895416056&imageWidth=898&imageHeight=700&lang=en&keyword=11SW-D/CR1055",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/C16&sd=&extent=840215.7647422651,807898.5730065152,840364.3093577347,808029.2026934859&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/C16",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/C17&sd=&extent=840211.1367190696,807924.101779626,840359.6813345392,808054.7314665967&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/C17",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/C20&sd=&extent=840113.5003587508,807808.998364021,840262.0449742205,807939.6280509917&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/C20",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/C21&sd=&extent=840053.4853482797,807805.7139604628,840202.0299637493,807936.3436474336&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/C21",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/C453&sd=&extent=839977.048320018,807655.3778521434,840125.5929354876,807786.0075391141&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/C453",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/C73&sd=&extent=840048.5596922651,807393.5641565144,840197.1043077348,807524.1938434851&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/C73",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/R85&sd=&extent=839974.5113211365,807532.4048523804,840123.0559366061,807663.0345393511&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/R85",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/C74&sd=&extent=839985.857442519,807504.6367132072,840134.4020579887,807635.266400178&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/C74",
                "https://www.slope.landsd.gov.hk/smris/generateSlopeReportBySlopeNo?sn=15NE-C/C201&sd=&extent=840076.7757046507,807344.8952674259,840225.3203201203,807475.5249543966&imageWidth=898&imageHeight=700&lang=en&keyword=15NE-C/C201",
            ];
            string folder = @"C:\Work\SlopeInspectionWebApi\SlopeInspectionWebApi\Resources";
            var random = new Random();
            using var client = new HttpClient();
            var retryPolicy = Policy
                .Handle<Exception>()
                .WaitAndRetryAsync(
                    3,
                    _ => TimeSpan.FromMilliseconds(random.Next(1000, 3000)),
                    (exception, retryDelay) =>
                    {
                        Console.WriteLine(exception.Message);
                    });
            var tasks = urls.Select(async url =>
            {
                await Task.Delay(random.Next(0, 1000));
                await retryPolicy.ExecuteAsync(async () => await Extract(client, folder, url));
            });
            await Task.WhenAll(tasks);
            Console.WriteLine($"Done in {sw.Elapsed.TotalSeconds}s.");
        }

        private static async Task Extract(HttpClient client, string folder, string url)
        {
            var uri = new Uri(url);
            var queries = uri.Query.TrimStart('?').Split('&').Select(q => q.Split('=')).ToDictionary(q => q[0], q => q[1]);
            string sn = queries["sn"];
            string extent = queries["extent"];
            string filePath = Path.Combine(folder, sn.Replace("-", "_").Replace("/", "_") + ".pdf");
            string fileFolder = Path.Combine(folder, Path.GetFileNameWithoutExtension(filePath));
            try
            {
                Directory.CreateDirectory(fileFolder);
                Console.WriteLine($"Loading PDF from {sn}...");
                var bytes = await client.GetByteArrayAsync(url);
                Console.WriteLine($"Extracting from {sn}...");
                using var pdf = new PdfDocument(bytes);
                var imageHelper = new PdfImageHelper();
                for (int i = 0; i < pdf.Pages.Count; i++)
                {
                    var page = pdf.Pages[i];
                    var imageInfos = imageHelper.GetImagesInfo(page);
                    for (int j = 0; j < imageInfos.Length; j++)
                    {
                        string fileName = $"Page{i + 1}-Image{j + 1}.png";
                        string imagePath = Path.Combine(fileFolder, fileName);
                        if (!File.Exists(imagePath))
                        {
                            using var fileStream = new FileStream(imagePath, FileMode.Create);
                            await imageInfos[j].Image.CopyToAsync(fileStream);
                        }
                    }
                }
                // Save the PDF document
                pdf.SaveToFile(filePath);
                // Save the extent
                string extentPath = Path.Combine(fileFolder, "extent.txt");
                if (!File.Exists(extentPath))
                {
                    await File.WriteAllTextAsync(extentPath, extent);
                }
                Console.WriteLine($"Extracted {sn} complete.");
            }
            catch (Exception)
            {
                Console.WriteLine($"Failed to extract from {sn}");
                throw;
            }
        }
    }
}
