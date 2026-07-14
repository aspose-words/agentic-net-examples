using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingFirstCharStyling
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
            // Register code page provider (required for some environments)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a foreach block that iterates over Items
            builder.Writeln("<<foreach [item in Items]>>");

            // Apply custom font styling to the first character of the Name:
            // - The first character is colored red using the textColor tag.
            // - The rest of the string is inserted without additional styling.
            builder.Writeln(
                "<<textColor [\"Red\"]>><<[item.Name.Substring(0,1)]>><</textColor>><<[item.Name.Substring(1)]>>");

            // End the foreach block
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Prepare sample data
            // -------------------------------------------------
            ReportModel model = new()
            {
                Items = new List<Item>
                {
                    new() { Name = "Apple" },
                    new() { Name = "Banana" },
                    new() { Name = "Cherry" },
                    new() { Name = "Date" }
                }
            };

            // -------------------------------------------------
            // 3. Load the template and build the report
            // -------------------------------------------------
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the model; the root object name is "model"
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 4. Save the generated report
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }
}
