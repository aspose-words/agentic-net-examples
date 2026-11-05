using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model classes
    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Quantity { get; set; }
    }

    public class Model
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            Model model = new()
            {
                Items = new()
                {
                    new Item { Name = "Apple", Quantity = 5 },
                    new Item { Name = "Banana", Quantity = 3 },
                    new Item { Name = "Cherry", Quantity = 12 }
                }
            };

            // Create template document with LINQ Reporting tags
            string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Items Report");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template for report generation
            Document reportDoc = new Document(templatePath);

            // Build the report using the model as the root object named "model"
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            // Indicate completion (no interactive input)
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
