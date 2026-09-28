using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingRemoveEmptyParagraphs
{
    // Sample data model
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for Aspose.Words)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create template document programmatically
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Static content
            builder.Writeln("Report Header");
            // Empty paragraph that should be removed
            builder.Writeln("");

            // Begin foreach loop over Items
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Name]>> - <<[item.Value]>>");
            // Empty paragraph inside the loop (should be removed)
            builder.Writeln("");
            builder.Writeln("<</foreach>>");

            // Footer
            builder.Writeln("Report Footer");
            // Save the template
            string templatePath = Path.Combine(outputDir, "Template.docx");
            template.Save(templatePath);

            // Load the template for reporting
            Document doc = new Document(templatePath);

            // Prepare sample data
            ReportModel model = new()
            {
                Items = new()
                {
                    new Item { Name = "Apple", Value = 10 },
                    new Item { Name = "Banana", Value = 20 },
                    new Item { Name = "Cherry", Value = 30 }
                }
            };

            // Configure ReportingEngine with RemoveEmptyParagraphs option
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            string outputPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(outputPath);
        }
    }
}
