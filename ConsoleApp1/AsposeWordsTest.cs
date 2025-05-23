using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace ConsoleApp1
{
    internal class AsposeWordsTest
    {
        private readonly License license = new License();
        private const int _otherNo = 25;

        class ModelItem
        {
            public string PhotoNo { get; set; } = "";
            public int LocationReference { get; set; }
            public int[] MaintenanceActionItemNos { get; set; } = [];
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

        public AsposeWordsTest()
        {
            license.SetLicense(@"C:\Work\ConsoleApp1\ConsoleApp1\TestWord2_out\Aspose.WordsProductFamily.lic");
        }

        public void Run()
        {
            string path = @"C:\Test\SlopeInspectionReporter\template.docx";
            string dataDir = @"C:\Work\ConsoleApp1\ConsoleApp1\TestWord2_out\";
            Stream sitePlanPicture = File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片1.png");
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
                    MaintenanceActionItemNos = [],
                    ImageStream = File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片2.jpg")
                },
                new ModelItem()
                {
                    PhotoNo = "222.jpg",
                    LocationReference = 2,
                    ItemType = ItemType.Defect,
                    MaintenanceActionItemNos = [1, 15],
                    ImageStream = File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片3.jpg")
                },
                new ModelItem()
                {
                    PhotoNo = "333.jpg",
                    LocationReference = 3,
                    ItemType = ItemType.Defect,
                    MaintenanceActionItemNos = [1, 2, 3],
                    ImageStream = File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片4.jpg")
                },
                new ModelItem()
                {
                    PhotoNo = "444.jpg",
                    LocationReference = 4,
                    ItemType = ItemType.Defect,
                    MaintenanceActionItemNos = [4, 5, 25],
                    ImageStream = File.OpenRead(@"C:\Test\SlopeInspectionReporter\图片5.jpg"),
                    Other = "Other text",
                },
            };

            var doc = new Document(path);
            foreach (var kv in replaceTexts)
            {
                doc.Range.Replace(kv.Key, kv.Value);
            }
            var builder = new DocumentBuilder(doc);
            // Get Shapes
            GroupShape? generalViewShape = null;
            GroupShape? defectShape = null;
            foreach (var cell in doc.GetChildNodes(NodeType.Cell, true).OfType<Cell>())
            {
                if (cell.GetText().Equals("\r\rGeneral View\a"))
                {
                    if (cell.FirstChild is Paragraph paragraph1 && paragraph1.FirstChild is GroupShape groupShape)
                    {
                        generalViewShape = groupShape;
                    }
                }
                if (cell.GetText().Equals("\r\rDefect \a"))
                {
                    if (cell.FirstChild is Paragraph paragraph1 && paragraph1.FirstChild is GroupShape groupShape)
                    {
                        defectShape = groupShape;
                    }
                }
            }
            // Set item checked
            foreach (var cell in doc.GetChildNodes(NodeType.Cell, true).OfType<Cell>())
            {
                foreach (var kv in _maintenanceActionItems)
                {
                    if (cell.Range.Text.Equals($"{kv.Key}\a"))
                    {
                        var items = modelItems.Where(n => n.MaintenanceActionItemNos.Contains(kv.Value)).Select(i => i.LocationReference).ToList();
                        if (items.Count != 0)
                        {
                            if (cell.NextCell is { } nextCell)
                            {
                                if (nextCell.LastParagraph.Runs.OfType<Run>().FirstOrDefault() is Run run)
                                {
                                    run.Font.Name = "Wingdings";
                                    run.Text = ((char)254).ToString();
                                }
                            }
                            if (cell.NextCell?.NextCell is { } nextNextCell)
                            {
                                var run = new Run(doc, string.Join(", ", items));
                                run.Font.Name = "Times New Roman";
                                run.Font.Size = 8;
                                nextNextCell.LastParagraph.Runs.Add(run);
                            }
                        }
                        break;
                    }
                }
            }
            // Set item checked 2
            foreach (var cell in doc.GetChildNodes(NodeType.Cell, true).OfType<Cell>())
            {
                if (cell.Range.Text.Equals("Immediate EI for Maintenance\a"))
                {
                    if (page3Info.ImmediateEIForMaintenance)
                    {
                        if (cell.NextCell is { } nextCell)
                        {
                            if (nextCell.LastParagraph.Runs.OfType<Run>().FirstOrDefault() is Run run)
                            {
                                run.Font.Name = "Wingdings";
                                run.Text = ((char)254).ToString();
                            }
                        }
                    }
                }
                else if (cell.Range.Text.Equals("Immediate arrangement for investigation and repair of buried water-carrying service\a"))
                {
                    if (page3Info.ImmediateArrangementForInvestigationAndRepairOfBuriedWaterCarryingService)
                    {
                        if (cell.NextCell is { } nextCell)
                        {
                            if (nextCell.LastParagraph.Runs.OfType<Run>().FirstOrDefault() is Run run)
                            {
                                run.Font.Name = "Wingdings";
                                run.Text = ((char)254).ToString();
                            }
                        }
                    }
                }
                else if (cell.Range.Text.Equals("In the country park / MTRC or KCRC protection zone:\a"))
                {
                    if (page3Info.InTheCountryParkMTRCOrKCRCProtectionZone)
                    {
                        if (cell.NextCell is { } nextCell)
                        {
                            if (nextCell.LastParagraph.Runs.OfType<Run>().FirstOrDefault() is Run run)
                            {
                                run.Font.Name = "Wingdings";
                                run.Text = ((char)254).ToString();
                            }
                        }
                    }
                }
                else if (cell.Range.Text.Equals("Access type (concrete / steel):\a"))
                {
                    if (cell.NextCell is { } nextCell)
                    {
                        nextCell.Range.Replace("{Access type}", page3Info.AccessType switch
                        {
                            AccessType.Concrete => "Concrete",
                            AccessType.Steel => "Steel",
                            _ => throw new NotImplementedException(),
                        });
                    }
                }
            }

            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>();
            // Weather
            if (paragraphs.FirstOrDefault(n => n.GetText().Equals("{Weathers}\a")) is Paragraph weathersParagraph)
            {
                var runs = new List<Run>();
                var weathers = Enum.GetValues<Weather>();
                Run? weatherRun = null;
                for (int i = 0; i < weathers.Length; i++)
                {
                    var item = weathers[i];
                    bool isChecked = page1Info.Weather == item;
                    var run = new Run(doc, page1Info.Weather.ToString());
                    run.Font.Name = "Times New Roman";
                    run.Font.Size = 10;
                    run.Text = Enum.GetName(item);
                    run.Font.StrikeThrough = !isChecked;
                    if (isChecked)
                    {
                        weatherRun = run;
                    }
                    runs.Add(run);
                    if (i != weathers.Length - 1)
                    {
                        var runSeparator = new Run(doc, page1Info.Weather.ToString());
                        runSeparator.Font.Name = "Times New Roman";
                        runSeparator.Font.Size = 10;
                        runSeparator.Text = " / ";
                        runSeparator.Font.StrikeThrough = true;
                        runs.Add(runSeparator);
                    }
                }
                if (weatherRun != null)
                {
                    var linkedList = new LinkedList<Run>(runs);
                    if (linkedList.Find(weatherRun) is { } node)
                    {
                        if (node.Previous is { } previous)
                        {
                            previous.Value.Font.StrikeThrough = false;
                        }
                        if (node.Next is { } next)
                        {
                            next.Value.Font.StrikeThrough = false;
                        }
                    }
                }
                weathersParagraph.RemoveAllChildren();
                foreach (var run in runs)
                {
                    weathersParagraph.AppendChild(run);
                }
            }
            // Site Plan
            if (generalViewShape != null && defectShape != null)
            {
                var sitePlanParagraph = paragraphs.FirstOrDefault(n => n.GetText().Equals("{Site Plan}\a"));
                if (sitePlanParagraph != null)
                {
                    builder.MoveTo(sitePlanParagraph);
                    sitePlanParagraph.RemoveAllChildren();
                    builder.InsertImage(sitePlanPicture);
                    foreach (var (item, index) in modelItems.Select((item, index) => (item, index)))
                    {
                        var clone = item.ItemType switch
                        {
                            ItemType.GeneralView => (GroupShape)generalViewShape.Clone(true),
                            ItemType.Defect => (GroupShape)defectShape.Clone(true),
                            _ => throw new NotImplementedException(),
                        };
                        SetShapeText(clone, doc, item.LocationReference.ToString());
                        clone.Left += (clone.Width) * index + 100;
                        clone.Top += (clone.Height) * index + 100;
                        clone.Rotation = 30 * index;
                        builder.InsertNode(clone);
                    }
                }
            }
            // foreach table
            var foreachStart = paragraphs.FirstOrDefault(n => n.GetText().Equals("{foreach}\r"));
            var foreachEnd = paragraphs.FirstOrDefault(n => n.GetText().Equals("{end}\r"));
            if (foreachStart != null && foreachEnd != null)
            {
                var nodes = GetNodesBetween(foreachStart, foreachEnd);
                foreachEnd.RemoveAllChildren();
                foreach (var item in modelItems)
                {
                    foreach (var node in nodes) // Table, only one node
                    {
                        var clone = node.Clone(true);
                        if (clone is CompositeNode compositeNode)
                        {
                            if (compositeNode.GetChildNodes(NodeType.Paragraph, true)
                                .OfType<Paragraph>()
                                .FirstOrDefault(p => p.GetText().Equals("{Photo}\a")) is { } photoParagraph)
                            {
                                if (item.ImageStream != null)
                                {
                                    builder.MoveTo(photoParagraph);
                                    photoParagraph.RemoveAllChildren();
                                    builder.InsertImage(item.ImageStream);
                                }
                            }
                            if (compositeNode.GetChildNodes(NodeType.Paragraph, true)
                                .OfType<Paragraph>()
                                .FirstOrDefault(p => p.GetText().Equals("{Maintenance Action Item}\a")) is { } maintenanceActionItemParagraph)
                            {
                                foreach (var text in GetItemTexts(item))
                                {
                                    var paragraph = maintenanceActionItemParagraph.Clone(true);
                                    paragraph.Range.Replace("{Maintenance Action Item}", text);
                                    maintenanceActionItemParagraph.ParentNode.InsertBefore(paragraph, maintenanceActionItemParagraph);
                                }
                                maintenanceActionItemParagraph.Remove();
                            }
                        }
                        clone.Range.Replace("{Photo No}", item.PhotoNo);
                        clone.Range.Replace("{location reference}", item.LocationReference.ToString());
                        foreachEnd.ParentNode.InsertBefore(clone, foreachEnd);
                    }
                    foreachEnd.ParentNode.InsertBefore(foreachEnd.Clone(true), foreachEnd);
                }
                // Clear template foreach
                foreachStart.Remove();
                foreachEnd.Remove();
                nodes.ForEach(n => n.Remove());
            }
            // Save
            var outputParameters = doc.Save(dataDir + "Shape_InsertShapeUsingDocumentBuilder_out.docx");
            string contentType = outputParameters.ContentType;
            Console.WriteLine(contentType);
        }
        private List<string> GetItemTexts(ModelItem item)
        {
            List<string> texts = item.ItemType switch
            {
                ItemType.GeneralView => ["General View"],
                ItemType.Defect => [.. item.MaintenanceActionItemNos
                                           .Where(i => i != _otherNo)
                                           .Select(n => _maintenanceActionItems.FirstOrDefault(m => m.Value == n).Key)],
                _ => throw new NotImplementedException()
            };
            if (item.MaintenanceActionItemNos.Contains(_otherNo) && !string.IsNullOrEmpty(item.Other))
            {
                texts.Add(item.Other);
            }
            return texts;
        }
        private static List<Node> GetNodesBetween(Paragraph foreachStart, Paragraph foreachEnd)
        {
            var nodes = new List<Node>();
            Node next = foreachStart;
            while (next.NextSibling != foreachEnd)
            {
                nodes.Add(next.NextSibling);
                next = next.NextSibling;
            }
            return nodes;
        }
        private static void SetShapeText(GroupShape groupShape, Document doc, string text)
        {
            if (groupShape.GetChildNodes(NodeType.Paragraph, true).FirstOrDefault() is Paragraph p)
            {
                p.AppendChild(new Run(doc, text));
            }
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
