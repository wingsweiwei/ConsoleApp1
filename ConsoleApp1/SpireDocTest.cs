using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System.Drawing;
using System.Linq;

namespace ConsoleApp1
{
    internal class SpireDocTest
    {
        public Document Run()
        {
            string path = @"C:\Test\SlopeInspectionReporter\template.docx";
            var dic = new Dictionary<string, string>()
            {
                { "{Date of Inspection}", "07 May 2025" },
                { "{Next RMI before}", "06 May 2026" },
            };
            var maintenanceActionItems = new Dictionary<string, int>()
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
            var maintenanceActionItems2 = new Dictionary<int, string>();
            var random1 = new Random();
            var random2 = new Random();
            foreach (var item in maintenanceActionItems.Values)
            {
                var array = new int[random1.Next(0, 3)];
                for (int i = 0; i < array.Length; i++)
                {
                    array[i] = random2.Next(1, 10);
                }
                maintenanceActionItems2.Add(item, string.Join(", ", array.Order()));
            }
            using Document doc = new Document(path, true);
            foreach (var section in doc.Sections.OfType<Section>())
            {
                foreach (var headerFooter in section.HeadersFooters.OfType<HeaderFooter>())
                {
                    //Console.WriteLine(headerFooter.Text);
                }
                foreach (var table in section.Tables.OfType<Table>())
                {
                    foreach (var row in table.Rows.OfType<TableRow>())
                    {
                        int actionRequireIndex = 0;
                        foreach (var (cell, index) in row.Cells.OfType<TableCell>().Select((c, i) => (c, i)))
                        {
                            foreach (var paragraph in cell.Paragraphs.OfType<Paragraph>())
                            {
                                foreach (var item in paragraph.Items)
                                {
                                    if (item is ShapeGroup shape)
                                    {
                                        var obj = shape.ChildObjects[0] as ShapeObject;
                                    }
                                    if (item is ShapeObject shapeObject)
                                    {
                                    }
                                    //Console.WriteLine(item);
                                }

                                if (!string.IsNullOrEmpty(paragraph.Text))
                                {
                                    ReplaceTest(dic, paragraph);
                                    if (index == 0)
                                    {
                                        actionRequireIndex = GetActionItemIndex(maintenanceActionItems, actionRequireIndex, paragraph);
                                    }
                                }
                                SetActionItem(maintenanceActionItems2, actionRequireIndex, index, paragraph);

                                if (paragraph.Text.Contains("{Site Plan}"))
                                {
                                    var shape1Cell = table.Rows.OfType<TableRow>().SelectMany(r => r.Cells.OfType<TableCell>()).FirstOrDefault(c => c.Paragraphs.OfType<Paragraph>().Any(p => p.Text.Contains("General View")));
                                    var g1 = shape1Cell.FirstChild.FirstChild;
                                    AddSitePlan(doc, paragraph);
                                    var clone = g1.Clone() as ShapeGroup;
                                    clone.Rotation = 90;
                                    ((Paragraph)clone.FirstChild.FirstChild).Text = "Yes";
                                    paragraph.ChildObjects.Add(clone);
                                    var clone2 = g1.Clone() as ShapeGroup;
                                    clone2.Rotation = 180;
                                    ((Paragraph)clone2.FirstChild.FirstChild).Text = "No";
                                    paragraph.ChildObjects.Add(clone2);
                                }
                            }
                        }
                    }
                }
            }
            doc.SaveToFile(@"C:\Test\SlopeInspectionReporter\output.docx");
            return doc;
        }
        private static int GetActionItemIndex(Dictionary<string, int> maintenanceActionItems, int actionRequireIndex, Paragraph paragraph)
        {
            if (maintenanceActionItems.TryGetValue(paragraph.Text, out var value))
            {
                actionRequireIndex = value;
            }

            return actionRequireIndex;
        }
        private static void SetActionItem(Dictionary<int, string> maintenanceActionItems2, int actionRequireIndex, int index, Paragraph paragraph)
        {
            if (actionRequireIndex != 0)
            {
                if (maintenanceActionItems2.TryGetValue(actionRequireIndex, out var text))
                {
                    if (!string.IsNullOrEmpty(text))
                    {
                        if (index == 1) // Action Require
                                        //if (paragraph.Text == "☐") // Action Require
                        {
                            paragraph.Items.Clear();
                            var symbol = paragraph.AppendSymbol(254);
                            symbol.FontName = "Wingdings";
                        }
                        if (index == 2) // Location Reference
                        {
                            paragraph.Text = text;
                        }
                    }
                }
            }
        }
        private static void ReplaceTest(Dictionary<string, string> dic, Paragraph paragraph)
        {
            foreach (var kv in dic)
            {
                if (paragraph.Text.Contains(kv.Key))
                {
                    paragraph.Text = paragraph.Text.Replace(kv.Key, kv.Value);
                }
            }
        }
        private static void AddSitePlan(Document doc, Paragraph paragraph)
        {
            paragraph.Items.Clear();
            paragraph.AppendPicture(@"C:\Test\SlopeInspectionReporter\图片1.png");
            var shape = paragraph.AppendShape(100, 100, ShapeType.Ellipse);
            shape.FillColor = Color.Blue;
            var group = paragraph.AppendShapeGroup(44.05f, 28.3f);
            //autoshape
            group.ChildObjects.Add(new ShapeObject(doc, ShapeType.StraightConnector1)
            {
                Width = 800,
                Height = 1000,
                FillColor = Color.White,
                StrokeColor = Color.Red,
                HorizontalPosition = 200,
                VerticalPosition = 500,
            });
            var shape2 = new ShapeObject(doc, ShapeType.Ellipse)
            {
                Width = 200,
                Height = 1000,
                FillColor = Color.White,
                StrokeColor = Color.Red,
            };
            //var p = new Paragraph(doc)
            //{
            //    Text = "test",
            //};
            //shape2.ChildObjects.Add(p);
            var clone = shape2.Clone();
            group.ChildObjects.Add(shape2);
            group.HorizontalPosition = 200;
            group.VerticalPosition = 20;
        }
    }
}
