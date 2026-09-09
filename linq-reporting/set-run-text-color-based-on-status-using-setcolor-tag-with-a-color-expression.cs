using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Status = "Success" },
                new Item { Status = "Failed" },
                new Item { Status = "Pending" }
            }
        };

        // Create the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Use textColor tag with a color expression based on the item's status.
        // The Color property of Item returns a color name string.
        builder.Writeln("<<textColor [item.Color]>>Status: <<[item.Status]>> <</textColor>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to a temporary file.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var loadedTemplate = new Document(templatePath);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        loadedTemplate.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model with a computed Color property.
public class Item
{
    public string Status { get; set; } = string.Empty;

    // Returns a color name based on the status.
    public string Color =>
        Status switch
        {
            "Success" => "Green",
            "Failed" => "Red",
            "Pending" => "Orange",
            _ => "Black"
        };
}
