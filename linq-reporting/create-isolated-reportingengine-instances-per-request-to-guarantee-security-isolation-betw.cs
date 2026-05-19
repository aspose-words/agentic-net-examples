using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineIsolationExample
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        public string UserName { get; set; } = string.Empty;
        public List<string> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // Step 1: Create a reusable template document containing LINQ Reporting tags.
            // -----------------------------------------------------------------
            const string templatePath = "template.docx";
            CreateTemplate(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Simulate two independent user requests.
            // Each request gets its own ReportingEngine instance and its own data.
            // -----------------------------------------------------------------
            var request1Data = new ReportModel
            {
                UserName = "Alice",
                Items = new() { "Item A", "Item B" }
            };
            var request2Data = new ReportModel
            {
                UserName = "Bob",
                Items = new() { "Item X", "Item Y", "Item Z" }
            };

            // Process first request.
            ProcessRequest(templatePath, "output1.docx", request1Data);

            // Process second request.
            ProcessRequest(templatePath, "output2.docx", request2Data);
        }

        // Creates the template file with LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Header showing the user name.
            builder.Writeln("Report for user: <<[model.UserName]>>");
            builder.Writeln();

            // Loop over the Items collection.
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in model.Items]>>");
            builder.Writeln("- <<[item]>>");
            builder.Writeln("<</foreach>>");

            // Save the template for later reuse.
            doc.Save(filePath);
        }

        // Loads the template, builds the report with isolated engine, and saves the result.
        private static void ProcessRequest(string templatePath, string outputPath, ReportModel data)
        {
            // Load the template document.
            var doc = new Document(templatePath);

            // Each request gets its own ReportingEngine instance.
            var engine = new ReportingEngine
            {
                // No special options required for this simple example.
                Options = ReportBuildOptions.None
            };

            // Build the report. The root object name must match the tag prefix ("model").
            bool success = engine.BuildReport(doc, data, "model");

            // In a real scenario you might check 'success' when InlineErrorMessages option is used.
            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
