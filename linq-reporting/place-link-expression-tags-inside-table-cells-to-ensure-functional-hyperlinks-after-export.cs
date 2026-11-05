using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqLinkInTable
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Sample data.
            ReportModel model = new()
            {
                Items = new()
                {
                    new Item { Url = "https://example.com/1", LinkText = "Example 1" },
                    new Item { Url = "https://example.com/2", LinkText = "Example 2" },
                    new Item { Url = "https://example.com/3", LinkText = "Example 3" }
                }
            };

            // Create the template document.
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            // Begin foreach block.
            builder.Writeln("<<foreach [item in Items]>>");

            // Start table inside the foreach block.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Link");
            builder.EndRow();

            // Data row for each item.
            builder.InsertCell();
            builder.Writeln("<<link [item.Url] [item.LinkText]>>");
            builder.EndRow();

            // End table.
            builder.EndTable();

            // End foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template.
            string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            Document reportDoc = new(templatePath);

            // Build the report.
            ReportingEngine engine = new();
            engine.Options = ReportBuildOptions.None;
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }

    // Root data model.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Item model representing a hyperlink.
    public class Item
    {
        public string Url { get; set; } = "";
        public string LinkText { get; set; } = "";
    }
}
