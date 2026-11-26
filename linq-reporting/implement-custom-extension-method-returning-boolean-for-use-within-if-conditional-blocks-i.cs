using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;

    // Computed property used by the template.
    public bool IsEven => Index % 2 == 0;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" },
                new Item { Index = 4, Name = "Delta" }
            }
        };

        // Create a template document.
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("LINQ Reporting Example with Custom Boolean Extension Method");
        builder.Writeln();

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Index with conditional background color.
        builder.Writeln(
            "<<if [item.IsEven]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Index]>> <</backColor>><</if>>" +
            "<<if [!item.IsEven]>>" +
            "<<[item.Index]>> " +
            "<</if>>");

        // Name with same conditional background color.
        builder.Writeln(
            "<<if [item.IsEven]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [!item.IsEven]>>" +
            "<<[item.Name]>> " +
            "<</if>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the output.
        const string outputPath = "output.docx";
        reportDoc.Save(outputPath);

        // Indicate result (no interactive input).
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
