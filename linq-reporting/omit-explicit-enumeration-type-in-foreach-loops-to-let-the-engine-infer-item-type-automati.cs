using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingInferenceExample
{
    // Data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new()
                {
                    new Item { Index = 1, Name = "Alpha" },
                    new Item { Index = 2, Name = "Beta" },
                    new Item { Index = 3, Name = "Gamma" }
                }
            };

            // Create the template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Report Items:");
            // Foreach tag without explicit type; the engine infers the item type.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(templatePath);

            // Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
