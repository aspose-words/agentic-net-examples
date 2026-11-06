using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin the foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start the table that will be repeated for each item.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row – will be generated for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();

        // Conditional background color: yellow when IsHighlighted is true.
        builder.Writeln(
            "<<if [item.IsHighlighted]>>" +
            "<<backColor [\"Yellow\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [!item.IsHighlighted]>>" +
            "<<[item.Name]>>" +
            "<</if>>");

        builder.EndRow();

        // End the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var templateDoc = new Document(templatePath);

        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Index = 1, Name = "Alice",   IsHighlighted = false },
                new Item { Index = 2, Name = "Bob",     IsHighlighted = true  },
                new Item { Index = 3, Name = "Charlie", IsHighlighted = false },
                new Item { Index = 4, Name = "Diana",   IsHighlighted = true  },
                new Item { Index = 5, Name = "Eve",     IsHighlighted = false }
            }
        };

        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        templateDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public bool IsHighlighted { get; set; }
}
