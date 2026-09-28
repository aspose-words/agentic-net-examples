using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingReflectionOptimization
{
    // Sample data model
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure code page provider is available (required by Aspose.Words for some encodings)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Create a large data set
            var model = new ReportModel();
            for (int i = 1; i <= 1000; i++)
            {
                model.Items.Add(new Item
                {
                    Index = i,
                    Name = $"Item #{i}"
                });
            }

            // Create a template document programmatically
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a simple heading
            builder.Writeln("Large Data Set Report");
            builder.Writeln();

            // Insert LINQ Reporting tags for iterating over Items
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Index: <<[item.Index]>> - Name: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to a temporary file
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            doc.Save(templatePath);

            // Load the template for reporting
            var templateDoc = new Document(templatePath);

            // Enable reflection optimization (caching) for the reporting engine
            ReportingEngine.UseReflectionOptimization = true;

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // Save the generated report
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report_Output.docx");
            templateDoc.Save(outputPath);

            // Optionally, indicate completion (no interactive input)
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
