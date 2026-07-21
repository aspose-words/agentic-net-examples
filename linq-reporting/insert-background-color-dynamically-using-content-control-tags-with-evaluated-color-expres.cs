using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Step 1: Create the LINQ Reporting template.
        const string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a simple title.
        builder.Writeln("Dynamic Background Colors Report");
        builder.Writeln();

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Insert a paragraph whose background color is taken from the data source.
        // The backColor tag evaluates the expression inside the brackets.
        builder.Writeln("<<backColor [item.Color]>> <<[item.Name]>> <</backColor>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Step 2: Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Step 3: Prepare the data model.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Color = "\"LightSalmon\"" },
                new Item { Name = "Banana", Color = "\"LightYellow\"" },
                new Item { Name = "Cherry", Color = "\"LightPink\"" },
                new Item { Name = "Date",   Color = "\"LightGray\"" }
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(reportDoc, model, "model");

        // Step 5: Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model referenced by the template as <<[model...]>>
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item containing a name and a background color expression.
public class Item
{
    // The name to display.
    public string Name { get; set; } = string.Empty;

    // Color expression returned as a quoted string (e.g., "\"LightGray\"").
    public string Color { get; set; } = string.Empty;
}
