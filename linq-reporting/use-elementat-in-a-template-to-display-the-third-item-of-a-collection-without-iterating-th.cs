using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingElementAtExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -------------------------------------------------
            // Create the template document programmatically
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Write a simple paragraph that uses ElementAt to fetch the third item (index 2)
            builder.Writeln("Third item name: <<[model.Items.ElementAt(2).Name]>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template for report generation
            // -------------------------------------------------
            var doc = new Document(templatePath);

            // -------------------------------------------------
            // Prepare sample data
            // -------------------------------------------------
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new() { Name = "Alpha" },
                    new() { Name = "Beta" },
                    new() { Name = "Gamma" },   // This is the third item (index 2)
                    new() { Name = "Delta" }
                }
            };

            // -------------------------------------------------
            // Build the report using LINQ Reporting Engine
            // -------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            doc.Save(outputPath);

            // Indicate completion (no interactive input)
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
