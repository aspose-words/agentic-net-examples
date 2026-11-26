using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingElementAtExample
{
    // Data model classes
    public class ReportModel
    {
        // Collection of items
        public List<Item> Items { get; set; } = new();

        // Fetch the second‑to‑last item using ElementAt and collection length
        public Item SecondToLastItem => Items.ElementAt(Items.Count - 2);
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output folder
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create the template document programmatically
            string templatePath = Path.Combine(outputDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Second to last item:");
            builder.Writeln("Name: <<[model.SecondToLastItem.Name]>>");
            builder.Writeln("Value: <<[model.SecondToLastItem.Value]>>");
            templateDoc.Save(templatePath);

            // Load the template for reporting
            Document doc = new Document(templatePath);

            // Prepare sample data
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Item A", Value = 10 },
                    new Item { Name = "Item B", Value = 20 },
                    new Item { Name = "Item C", Value = 30 },
                    new Item { Name = "Item D", Value = 40 }
                }
            };

            // Build the report using Aspose.Words LINQ Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            string outputPath = Path.Combine(outputDir, "output.docx");
            doc.Save(outputPath);
        }
    }
}
