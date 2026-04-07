using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Name = "Apple",  Color = "Red" },
                new Item { Name = "Banana", Color = "Green" },
                new Item { Name = "Grape",  Color = "#800080" } // Purple using hex code.
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }

    // Creates a simple Word template containing a foreach loop and a textColor tag.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Apply dynamic text color based on the item's Color property.
        builder.Writeln("<<textColor [item.Color]>>Item: <<[item.Name]>> <</textColor>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Wrapper class for the data source.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple data item with a name and a color expression.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Color { get; set; } = string.Empty;
}
