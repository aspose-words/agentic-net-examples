using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create the LINQ Reporting template.
        const string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Apply numbering to each list item.
        builder.ListFormat.ApplyNumberDefault();
        // Insert a bookmark around the item title.
        builder.Writeln("<<bookmark [item.BookmarkName]>><<[item.Title]>><</bookmark>>");
        // Remove numbering formatting for subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Title = "First item", BookmarkName = "bm1" },
                new Item { Title = "Second item", BookmarkName = "bm2" },
                new Item { Title = "Third item", BookmarkName = "bm3" }
            }
        };

        // Build the report using LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }

    // Root data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Item model representing each list entry.
    public class Item
    {
        public string Title { get; set; } = "";
        public string BookmarkName { get; set; } = "";
    }
}
