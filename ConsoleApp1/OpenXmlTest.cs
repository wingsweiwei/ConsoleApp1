using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;

namespace ConsoleApp1
{
    internal class OpenXmlTest
    {
        private const int _otherNo = 25;

        class ModelItem
        {
            public string PhotoNo { get; set; } = "";
            public int LocationReference { get; set; }
            public int[] MaintenanceActionItemNos { get; set; } = Array.Empty<int>();
            public string Other { get; set; } = "";
            public ItemType ItemType { get; set; }
            public Stream? ImageStream { get; set; }
        }
        class PageHeaderInfo
        {
            public string InspectionNo { get; set; } = "";
            public string HyDReferenceNo { get; set; } = "";
            public string SlopeFileReference { get; set; } = "";
            public string SlopeRetainingWallReferenceNo { get; set; } = "";
            public string Location { get; set; } = "";
        }
        class Page1Info
        {
            public string DateOfInspection { get; set; } = "";
            public string NextRMIBefore { get; set; } = "";
            public Weather Weather { get; set; }
            public string WorksOrderNo { get; set; } = "";
            public string RecommendedDateForCompletionOfAboveWorks { get; set; } = "";
            public string ActualCompletionDateOfAllAboveWorks { get; set; } = "";
        }
        class Page3Info
        {
            public bool ImmediateEIForMaintenance { get; set; }
            public bool ImmediateArrangementForInvestigationAndRepairOfBuriedWaterCarryingService { get; set; }
            public bool InTheCountryParkMTRCOrKCRCProtectionZone { get; set; }
            public AccessType AccessType { get; set; }
            public string OtherObservations { get; set; } = "";
            public string InspectedByCheckedBy { get; set; } = "";
            public string InspectedOfCheckedOf { get; set; } = "";
            public string Signature1 { get; set; } = "";
            public string Date1 { get; set; } = "";
            public string DateOfNextInspection1 { get; set; } = "";
            public string MaintenanceWorksFinishedDate1 { get; set; } = "";
            public string ReceivedBy { get; set; } = "";
            public string ReceivedOf { get; set; } = "";
            public string Signature2 { get; set; } = "";
            public string Date2 { get; set; } = "";
            public string Remark { get; set; } = "";
        }
        class Page5Info
        {
            public string MaintenanceWorksArrangedBy { get; set; } = "";
            public string MaintenanceWorksArrangedOf { get; set; } = "";
            public string Signature3 { get; set; } = "";
            public string Date3 { get; set; } = "";
            public string DateOfNextInspection3 { get; set; } = "";
            public string MaintenanceWorksFinishedDate3 { get; set; } = "";
            public string MaintenanceWorksCarriedOutBy { get; set; } = "";
            public string MaintenanceWorksCarriedOutOf { get; set; } = "";
            public string Signature4 { get; set; } = "";
            public string Date4 { get; set; } = "";
            public string MaintenanceWorksCarriedOutOn { get; set; } = "";
        }
        enum ItemType
        {
            GeneralView,
            Defect,
        }
        enum Weather
        {
            Fine,
            Cloudy,
            Showery,
            Rainy,
        }
        enum AccessType
        {
            Concrete,
            Steel,
        }
        public void Run()
        {
            Stopwatch sw = Stopwatch.StartNew();
            string path = @"C:\Test\SlopeInspectionReporter\template.docx";
            string outputPath = @"C:\Test\SlopeInspectionReporter\output.docx";
            using var ms = new MemoryStream();
            using (var templateStream = File.OpenRead(path))
            {
                templateStream.CopyTo(ms);
            }
            Stream sitePlanPicture = File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片1.png");
            Stream[] images = new Stream[]
            {
                File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片2.jpg"),
                File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片3.jpg"),
                File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片4.jpg"),
                File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片5.jpg"),
            };

            var replaceTexts = new Dictionary<string, string>();
            var pageHeaderInfo = new PageHeaderInfo()
            {
                InspectionNo = "RMI-12403C00710-20250507-SW01",
                HyDReferenceNo = "12403C00710",
                SlopeFileReference = "",
                SlopeRetainingWallReferenceNo = "11SW-C/C176",
                Location = "WITHIN IL623S.13,RP & S.14,RP AND GOVERNMENT LAND TO ITS WEST, ABUTTING VICTORIA ROAD",
            };
            replaceTexts.Add("{Inspection No}", pageHeaderInfo.InspectionNo);
            replaceTexts.Add("{HyD Reference No}", pageHeaderInfo.HyDReferenceNo);
            replaceTexts.Add("{Slope File Reference}", pageHeaderInfo.SlopeFileReference);
            replaceTexts.Add("{Slope / Retaining Wall Reference No}", pageHeaderInfo.SlopeRetainingWallReferenceNo);
            replaceTexts.Add("{Location}", pageHeaderInfo.Location);
            var page1Info = new Page1Info()
            {
                DateOfInspection = "07 May 2025",
                NextRMIBefore = "06 May 2026",
                Weather = Weather.Showery,
                WorksOrderNo = "Works Order No",
                RecommendedDateForCompletionOfAboveWorks = "06 June 2025",
                ActualCompletionDateOfAllAboveWorks = "Actual Completion",
            };
            replaceTexts.Add("{Date of Inspection}", page1Info.DateOfInspection);
            replaceTexts.Add("{Next RMI before}", page1Info.NextRMIBefore);
            replaceTexts.Add("{Works Order No}", page1Info.WorksOrderNo);
            replaceTexts.Add("{Recommended Date for Completion of Above Works}", page1Info.RecommendedDateForCompletionOfAboveWorks);
            replaceTexts.Add("{Actual Completion Date of All Above Works}", page1Info.ActualCompletionDateOfAllAboveWorks);
            var page3Info = new Page3Info()
            {
                ImmediateEIForMaintenance = true,
                ImmediateArrangementForInvestigationAndRepairOfBuriedWaterCarryingService = true,
                InTheCountryParkMTRCOrKCRCProtectionZone = true,
                AccessType = AccessType.Steel,
                OtherObservations = "Other Observations",
                InspectedByCheckedBy = "Inspected by",
                InspectedOfCheckedOf = "Inspected of",
                Signature1 = "Signature1",
                Date1 = "Date1",
                DateOfNextInspection1 = "Date of1",
                MaintenanceWorksFinishedDate1 = "Maintenance works1",
                ReceivedBy = "Received by",
                ReceivedOf = "Received of",
                Signature2 = "Signature2",
                Date2 = "Date2",
                Remark = "Remark",
            };
            replaceTexts.Add("{Other Observations}", page3Info.OtherObservations);
            replaceTexts.Add("{Inspected by / Checked by}", page3Info.InspectedByCheckedBy);
            replaceTexts.Add("{Inspected of / Checked of}", page3Info.InspectedOfCheckedOf);
            replaceTexts.Add("{Signature1}", page3Info.Signature1);
            replaceTexts.Add("{Date1}", page3Info.Date1);
            replaceTexts.Add("{Date of next inspection1}", page3Info.DateOfNextInspection1);
            replaceTexts.Add("{Maintenance works finished date1}", page3Info.MaintenanceWorksFinishedDate1);
            replaceTexts.Add("{Received by}", page3Info.ReceivedBy);
            replaceTexts.Add("{Received of}", page3Info.ReceivedOf);
            replaceTexts.Add("{Signature2}", page3Info.Signature2);
            replaceTexts.Add("{Date2}", page3Info.Date2);
            replaceTexts.Add("{Remark}", page3Info.Remark);
            var page5Info = new Page5Info()
            {
                MaintenanceWorksArrangedBy = "arrange by",
                MaintenanceWorksArrangedOf = "arrange of",
                Signature3 = "Signature3",
                Date3 = "Date3",
                DateOfNextInspection3 = "Date of next inspection3",
                MaintenanceWorksFinishedDate3 = "Maintenance date3",
                MaintenanceWorksCarriedOutBy = "carried by",
                MaintenanceWorksCarriedOutOf = "carried of",
                Signature4 = "Signature4",
                Date4 = "Date4",
                MaintenanceWorksCarriedOutOn = "carried on",
            };
            replaceTexts.Add("{Maintenance works arranged by}", page5Info.MaintenanceWorksArrangedBy);
            replaceTexts.Add("{Maintenance works arranged of}", page5Info.MaintenanceWorksArrangedOf);
            replaceTexts.Add("{Signature3}", page5Info.Signature3);
            replaceTexts.Add("{Date3}", page5Info.Date3);
            replaceTexts.Add("{Date of next inspection3}", page5Info.DateOfNextInspection3);
            replaceTexts.Add("{Maintenance works finished date3}", page5Info.MaintenanceWorksFinishedDate3);
            replaceTexts.Add("{Maintenance works carried out by}", page5Info.MaintenanceWorksCarriedOutBy);
            replaceTexts.Add("{Maintenance works carried out of}", page5Info.MaintenanceWorksCarriedOutOf);
            replaceTexts.Add("{Signature4}", page5Info.Signature4);
            replaceTexts.Add("{Date4}", page5Info.Date4);
            replaceTexts.Add("{Maintenance works carried out on}", page5Info.MaintenanceWorksCarriedOutOn);
            var modelItems = new List<ModelItem>
            {
                new ModelItem()
                {
                    PhotoNo = "111.jpg",
                    LocationReference = 1,
                    ItemType = ItemType.GeneralView,
                    MaintenanceActionItemNos = Array.Empty<int>(),
                    ImageStream = images[0]
                },
                new ModelItem()
                {
                    PhotoNo = "222.jpg",
                    LocationReference = 2,
                    ItemType = ItemType.Defect,
                    MaintenanceActionItemNos = [1, 15],
                    ImageStream = images[1]
                },
                new ModelItem()
                {
                    PhotoNo = "333.jpg",
                    LocationReference = 3,
                    ItemType = ItemType.Defect,
                    MaintenanceActionItemNos = [1, 2, 3],
                    ImageStream = images[2]
                },
                new ModelItem()
                {
                    PhotoNo = "444.jpg",
                    LocationReference = 4,
                    ItemType = ItemType.Defect,
                    MaintenanceActionItemNos = [4, 5, 25],
                    ImageStream = images[3],
                    Other = "Other text",
                },
            };

            using (WordprocessingDocument doc = WordprocessingDocument.Open(ms, true))
            {
                // Replace Text
                NewMethod(sitePlanPicture, replaceTexts, page1Info, modelItems, doc);
                //doc.MainDocumentPart.Document.Save();
            }
            using (var fs = File.Create(outputPath))
            {
                ms.Position = 0;
                ms.CopyTo(fs);
            }
            Console.WriteLine(sw.ElapsedMilliseconds);
        }

        private void NewMethod(Stream sitePlanPicture, Dictionary<string, string> replaceTexts, Page1Info page1Info, List<ModelItem> modelItems, WordprocessingDocument doc)
        {
            foreach (var kv in replaceTexts)
            {
                foreach (var part in doc.MainDocumentPart.HeaderParts)
                {
                    ReplaceText(part.Header, kv.Key, kv.Value);
                }
                ReplaceText(doc.MainDocumentPart.Document, kv.Key, kv.Value);
                foreach (var part in doc.MainDocumentPart.FooterParts)
                {
                    ReplaceText(part.Footer, kv.Key, kv.Value);
                }
            }

            // Weather
            var weathersParagraph = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>().FirstOrDefault(n => n.InnerText.Contains("{Weathers}"));
            if (weathersParagraph != null)
            {
                weathersParagraph.RemoveAllChildren();
                var weathers = Enum.GetValues<Weather>();
                for (int i = 0; i < weathers.Length; i++)
                {
                    var item = weathers[i];
                    bool isChecked = page1Info.Weather == item;
                    var run = new Run(new Text(Enum.GetName(item) ?? ""));
                    run.RunProperties = new RunProperties(new Bold(), new FontSize() { Val = "22" });
                    run.RunProperties.Append(new Strike() { Val = isChecked });
                    weathersParagraph.Append(run);
                    if (i != weathers.Length - 1)
                    {
                        var runSeparator = new Run(new Text(" / "));
                        runSeparator.RunProperties = new RunProperties(new Bold(), new FontSize() { Val = "22" });
                        runSeparator.RunProperties.Append(new Strike() { Val = true });
                        weathersParagraph.Append(runSeparator);
                    }
                }
            }

            // Site Plan
            var sitePlanParagraph = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>().FirstOrDefault(n => n.InnerText.Equals("{Site Plan}"));
            if (sitePlanParagraph != null)
            {
                sitePlanParagraph.RemoveAllChildren<Run>();
                var generalViewParagraph = doc.MainDocumentPart.Document.Body.Descendants<TableCell>().FirstOrDefault(n => n.InnerText.Contains("General View"));
                var defectParagraph = doc.MainDocumentPart.Document.Body.Descendants<TableCell>().FirstOrDefault(n => n.InnerText.Contains("Defect "));
                if (generalViewParagraph != null && defectParagraph != null)
                {
                    var clone = generalViewParagraph.GetFirstChild<Paragraph>().LastChild.CloneNode(true);
                    sitePlanParagraph.Append(clone);
                }
                InsertAPicture(doc.MainDocumentPart, sitePlanParagraph, sitePlanPicture);
                foreach (var (item, index) in modelItems.Select((item, index) => (item, index)))
                {
                    InsertText(doc.MainDocumentPart.Document.Body, item.LocationReference.ToString());
                }
            }

            // Foreach table
            var foreachStart = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().FirstOrDefault(n => n.InnerText.Contains("{foreach}"));
            var foreachEnd = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().FirstOrDefault(n => n.InnerText.Contains("{end}"));
            if (foreachStart != null && foreachEnd != null)
            {
                var nodes = GetNodesBetween(foreachStart, foreachEnd);
                foreachEnd.RemoveAllChildren();
                foreach (var item in modelItems)
                {
                    foreach (var node in nodes)
                    {
                        var clone = node.CloneNode(true);
                        if (clone is Table table)
                        {
                            foreach (var cell in table.Elements<TableRow>().SelectMany(r => r.Elements<TableCell>()))
                            {
                                foreach (var paragraph in cell.Elements<Paragraph>())
                                {
                                    foreach (var run in paragraph.Elements<Run>())
                                    {
                                        if (run.InnerText.Contains("{Photo}"))
                                        {
                                            if (item.ImageStream != null)
                                            {
                                                InsertAPicture(doc.MainDocumentPart, cell, item.ImageStream);
                                            }
                                        }
                                        else if (run.InnerText.Contains("{Maintenance Action Item}"))
                                        {
                                            foreach (var text in GetItemTexts(item))
                                            {
                                                var cloneParagraph = run.CloneNode(true) as Run;
                                                if (cloneParagraph != null)
                                                {
                                                    cloneParagraph.AppendChild(new Text(text));
                                                    paragraph.Append(cloneParagraph);
                                                }
                                            }
                                        }
                                        else if (run.InnerText.Contains("{Photo No}"))
                                        {
                                            run.AppendChild(new Text(item.PhotoNo));
                                        }
                                        else if (run.InnerText.Contains("{location reference}"))
                                        {
                                            run.AppendChild(new Text(item.LocationReference.ToString()));
                                        }
                                    }
                                }
                            }
                            foreachEnd.Parent.InsertBefore(table, foreachEnd);
                        }
                    }
                    foreachEnd.Parent.InsertBefore(foreachEnd.CloneNode(true), foreachEnd);
                }
                // Clear template foreach
                foreachStart.Remove();
                foreachEnd.Remove();
                nodes.ForEach(n => n.Remove());
            }
        }

        private static void ReplaceText(OpenXmlElement document, string searchText, string replaceText)
        {
            foreach (var paragraph in document.Descendants<Paragraph>())
            {
                if (paragraph.InnerText.Contains(searchText))
                {
                    var run = paragraph.GetFirstChild<Run>();
                    run.GetFirstChild<Text>().Text = paragraph.InnerText.Replace(searchText, replaceText);
                    paragraph.RemoveAllChildren<Run>();
                    paragraph.Append(run);
                }
            }
        }
        private static void InsertText(OpenXmlElement parent, string text)
        {
            parent.Append(new Paragraph(new Run(new Text(text))));
        }
        static void InsertAPicture(MainDocumentPart mainPart, OpenXmlElement parent, Stream stream)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            imagePart.FeedData(stream);
            AddImageToBody(parent, mainPart.GetIdOfPart(imagePart));
        }
        static void AddImageToBody(OpenXmlElement parent, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            parent.AppendChild(element);
        }

        private List<string> GetItemTexts(ModelItem item)
        {
            List<string> texts = item.ItemType switch
            {
                ItemType.GeneralView => ["General View"],
                ItemType.Defect => [.. item.MaintenanceActionItemNos
                                           .Where(i => i != _otherNo)
                                           .Select(n => _maintenanceActionItems.FirstOrDefault(m => m.Value == n).Key)],
                _ => throw new NotImplementedException(),
            };
            if (item.MaintenanceActionItemNos.Contains(_otherNo) && !string.IsNullOrEmpty(item.Other))
            {
                texts.Add(item.Other);
            }
            return texts;
        }
        private static List<OpenXmlElement> GetNodesBetween(Paragraph foreachStart, Paragraph foreachEnd)
        {
            var nodes = new List<OpenXmlElement>();
            OpenXmlElement next = foreachStart;
            while (next.NextSibling() != foreachEnd)
            {
                nodes.Add(next.NextSibling());
                next = next.NextSibling();
            }
            return nodes;
        }

        readonly Dictionary<string, int> _maintenanceActionItems = new Dictionary<string, int>()
        {
            { "Clear drainage channels of accumulated debris", 1},
            { "Repair cracked / damaged drainage channels or pavements along crest and toe of slope or retaining wall", 2},
            { "Repair or replace cracked or damaged impermeable slope surface cover", 3},
            { "Repair or replace damaged tree rings", 4},
            { "Repair or reprovision of damaged erosion control mats and / or wire mesh", 5 },
            { "Remove surface debris and vegetation that has caused severe cracking of slope surface cover and drainage channels", 6},
            { "Remove loose rock debris and undesirable vegetation from rock slopes or boulders", 7},
            { "Re-vegetate bare soil slope surface", 8},
            { "Repair pointings in masonry walls", 9},
            { "Unblock weepholes and outlet drainpipes", 10},
            { "Repair leaky exposed water-carrying services", 11},
            { "Repair or replace rusted slope furniture (e.g. Steel gates, boundary fences and stairs)", 12},
            { "Remove debris from defence measures", 13},
            { "Replace (including supply) defective / missing covers / gratings / frames / inlet traps", 14},
            { "Cutting / removal of any vegetation affecting the free flow of water in the drainage channels", 15},
            { "Cutting of grass", 16},
            { "Pruning of shrubs and trees on slopes", 17},
            { "Removal of fallen trees and broken tree branches", 18},
            { "Repair or replace of rusted, damaged or missing rock mesh (all types) on rock slopes", 19},
            { "Removal of stagnant water", 20},
            { "Removal / covering up of graffito on roadside slopes", 21},
            { "Repainting of rusted parts of boulder fences, steel elements of flexible barriers and steel baffles", 22},
            { "Conduct cutting operation of wire mesh on trees", 23},
            { "Apply chemical treatments such as chronic anti-coagulants rodenticides for rodent onto", 24},
            { "Others (specify works and give details)", 25},
        };
    }
}
