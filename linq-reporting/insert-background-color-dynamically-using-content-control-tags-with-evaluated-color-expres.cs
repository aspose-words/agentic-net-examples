using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class ReportModel
{
    // Collection of items to iterate over in the template.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    // Name of the item.
    public string Name { get; set; } = string.Empty;

    // Background color for the item (color name or HTML code).
    public string Color { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Write a foreach block that iterates over Items.
        // For each item we apply a backColor tag whose color is taken from the data source.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<backColor [item.Color]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</backColor>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Load the template back (required by the workflow).
        // -----------------------------------------------------------------
        var document = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Color = "LightYellow" },
                new Item { Name = "Banana", Color = "#FFFACD" }, // Light goldenrod yellow
                new Item { Name = "Cherry", Color = "LightCoral" }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        engine.BuildReport(document, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        document.Save(outputPath, SaveFormat.Docx);
    }
}
