using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    // Simple data model for list items.
    public class Item
    {
        public string Name { get; set; } = "";
        public string BookmarkName { get; set; } = "";
    }

    // Wrapper class required for LINQ Reporting root object.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a numbered list style.
        List list = templateDoc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // Insert LINQ Reporting tags.
        // The foreach tag iterates over the Items collection.
        // Inside each iteration we place a bookmark whose name comes from the data source.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</foreach>>");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "First item", BookmarkName = "bmFirst" },
                new Item { Name = "Second item", BookmarkName = "bmSecond" },
                new Item { Name = "Third item", BookmarkName = "bmThird" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        doc.Save(reportPath);
    }
}
