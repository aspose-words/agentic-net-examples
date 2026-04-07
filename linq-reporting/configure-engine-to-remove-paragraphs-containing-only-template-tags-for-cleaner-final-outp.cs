using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string? Optional { get; set; } // May be null, resulting in an empty paragraph.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -------------------------------------------------
        // 1. Create the template document with LINQ tags.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Static title.
        builder.Writeln("=== Report ===");

        // Dynamic title from the model.
        builder.Writeln("<<[model.Title]>>");

        // This paragraph may become empty after processing because Optional can be null.
        builder.Writeln("<<[model.Optional]>>");

        // Header for the items list.
        builder.Writeln("Items:");

        // Loop over the collection.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        ReportModel model = new()
        {
            Title = "Sample Report",
            Optional = null, // This will cause the paragraph to be empty.
            Items = new List<Item>
            {
                new Item { Name = "Apples",  Value = "10" },
                new Item { Name = "Oranges", Value = "5"  },
                new Item { Name = "Bananas", Value = "7"  }
            }
        };

        // -------------------------------------------------
        // 3. Configure the ReportingEngine to remove empty paragraphs.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 4. Save the final report.
        // -------------------------------------------------
        doc.Save(reportPath);
    }
}
