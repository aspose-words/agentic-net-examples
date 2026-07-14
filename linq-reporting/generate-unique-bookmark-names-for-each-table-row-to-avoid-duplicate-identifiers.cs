using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table class

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin the foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table that will be repeated for each item.
        Table table = builder.StartTable();

        // First (and only) cell of the row.
        builder.InsertCell();

        // Insert a bookmark whose name comes from the data source.
        // The bookmark wraps the displayed item name.
        builder.Writeln("<<bookmark [item.Bookmark]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data model with unique bookmark names.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<Item>()
        };

        // Sample data – each item gets a distinct bookmark name.
        for (int i = 1; i <= 5; i++)
        {
            model.Items.Add(new Item
            {
                Name = $"Item {i}",
                Bookmark = $"Row_{i}"
            });
        }

        // -----------------------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options

        // The root object name in the template is "model".
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties are initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Bookmark { get; set; } = string.Empty;
}
