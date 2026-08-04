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

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Use the textColor tag. The color expression is taken from the item's Color property.
        // The content inside the tag will be displayed in the chosen color.
        builder.Writeln("<<textColor [item.Color]>>Status: <<[item.Status]>> <</textColor>>");

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the data source.
        engine.BuildReport(doc, model);

        // Save the generated report.
        const string reportPath = "report.docx";
        doc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Status { get; set; } = string.Empty;

    // Returns a color name based on the status.
    public string Color =>
        Status switch
        {
            "Success" => "Green",
            "Failed"  => "Red",
            _         => "Gray"
        };
}
