using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data model.
        var model = new ReportModel
        {
            Name = "John Doe",
            Age = 30,
            Items = new List<Item>
            {
                new Item { Index = 1, Description = "Item One" },
                new Item { Index = 2, Description = "Item Two" }
            }
        };

        // Create a blank document and a builder to insert the template.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Use the default LINQ Reporting delimiters (<< and >>) for tags.
        builder.Writeln("Report for: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln(" - <<[item.Index]>>: <<[item.Description]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the model as the root data source named "model".
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Description { get; set; } = string.Empty;
}
