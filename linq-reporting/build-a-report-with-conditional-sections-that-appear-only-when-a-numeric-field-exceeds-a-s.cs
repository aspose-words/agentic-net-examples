using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Item
    {
        public string Name { get; set; } = "";
        public double Amount { get; set; }
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "template.docx";
            string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Sales Report");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Name]>>");
            // Conditional section: appears only when Amount exceeds 1000.
            builder.Writeln("<<if [item.Amount > 1000]>>");
            builder.Writeln("  High value sale: <<[item.Amount]>>");
            builder.Writeln("<</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template document for reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Product A", Amount = 750 },
                    new Item { Name = "Product B", Amount = 1250 },
                    new Item { Name = "Product C", Amount = 500 },
                    new Item { Name = "Product D", Amount = 2000 }
                }
            };

            // -----------------------------------------------------------------
            // Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may result from conditional sections.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
