using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        // Note: No property named MissingProperty – it will be treated as null.
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for older encodings if needed.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output documents.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -------------------------------------------------
            // Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a foreach loop that iterates over Items.
            // The second expression references a missing member to demonstrate AllowMissingMembers.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Name: <<[item.Name]>>");
            builder.Writeln("Missing: <<[item.MissingProperty]>>"); // This member does not exist.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template document for reporting.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // Prepare sample data.
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Alpha" },
                    new Item { Name = "Beta" },
                    new Item { Name = "Gamma" }
                }
            };

            // Configure the ReportingEngine to treat missing members as null.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.MissingMemberMessage = ""; // Empty string will result in blank output for missing members.

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
