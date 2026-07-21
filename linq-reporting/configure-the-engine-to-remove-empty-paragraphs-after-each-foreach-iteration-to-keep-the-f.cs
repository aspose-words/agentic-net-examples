using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data – some items have empty names to generate empty paragraphs.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alice" },
                new Item { Index = 2, Name = "" },          // Will produce an empty paragraph.
                new Item { Index = 3, Name = "Charlie" },
                new Item { Index = 4, Name = "" }           // Will produce an empty paragraph.
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting foreach block – each iteration writes the item's name.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Name]>>"); // Paragraph becomes empty when Name is empty.
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to remove empty paragraphs.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
