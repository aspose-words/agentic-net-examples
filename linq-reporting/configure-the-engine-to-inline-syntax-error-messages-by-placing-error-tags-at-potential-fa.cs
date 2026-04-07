using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string CustomerName { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        public List<Item> Items { get; set; } = new();
    }

    // Item class used in a collection.
    public class Item
    {
        public string Name { get; set; } = "Sample Item";
        public double Price { get; set; } = 9.99;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a blank document and build the template with LINQ Reporting tags.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Valid tags.
            builder.Writeln("Customer: <<[model.CustomerName]>>");
            builder.Writeln("Age: <<[model.Age]>>");

            // Tag that references a non‑existent member – will trigger an inline error message.
            builder.Writeln("Missing property: <<[model.Unknown]>>");

            // Foreach loop over a valid collection.
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before building the report).
            string templatePath = Path.Combine(outputDir, "Template.docx");
            templateDoc.Save(templatePath);

            // Load the template back.
            var loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            var model = new ReportModel
            {
                // Populate the collection with a couple of items.
                Items = new List<Item>
                {
                    new() { Name = "Apple", Price = 1.23 },
                    new() { Name = "Banana", Price = 0.99 }
                }
            };

            // Configure the reporting engine to inline error messages.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // Build the report. The returned flag indicates whether parsing succeeded.
            bool success = engine.BuildReport(loadedTemplate, model, "model");

            // Save the resulting document.
            string resultPath = Path.Combine(outputDir, "Report.docx");
            loadedTemplate.Save(resultPath);

            // Output simple status (no interactive prompts).
            Console.WriteLine($"Report generation success flag: {success}");
            Console.WriteLine($"Report saved to: {resultPath}");
        }
    }
}
