using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        Directory.CreateDirectory("output");

        // Step 1: Create the LINQ Reporting template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Apply dynamic text color based on the item's Color property.
        builder.Writeln("<<textColor [item.Color]>>Item: <<[item.Name]>> <</textColor>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine("output", "template.docx");
        template.Save(templatePath);

        // Step 2: Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Step 3: Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple",  Color = "Red" },
                new Item { Name = "Banana", Color = "Yellow" },
                new Item { Name = "Grape",  Color = "#800080" } // Purple using hex code.
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Step 5: Save the generated report.
        string outputPath = Path.Combine("output", "report.docx");
        reportDoc.Save(outputPath);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item with a name and a color expression.
public class Item
{
    public string Name { get; set; } = "";
    public string Color { get; set; } = "";
}
