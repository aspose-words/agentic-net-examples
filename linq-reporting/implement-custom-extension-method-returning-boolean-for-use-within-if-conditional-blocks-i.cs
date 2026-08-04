using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyExtensions
{
    // Extension method to check if an integer is even.
    public static bool IsEven(this int value) => value % 2 == 0;
}

// Data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" },
                new Item { Index = 4, Name = "Delta" }
            }
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Report Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        // Use the custom extension method inside an if tag.
        builder.Writeln("<<if [MyExtensions.IsEven(item.Index)]>>");
        builder.Writeln("<<backColor [\"LightGray\"]>><<[item.Name]>> (Even)<</backColor>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Register the static class containing the extension method.
        engine.KnownTypes.Add(typeof(MyExtensions));

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);
    }
}
