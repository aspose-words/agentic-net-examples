using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        // Start a data band that iterates over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Output the item name.
        builder.Writeln("Item: <<[item.Name]>>");
        // Force movement to the next item using a true condition.
        builder.Writeln("<<if [true]>><<next>>><</if>>");
        // End the data band.
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha" },
                new Item { Name = "Beta" },
                new Item { Name = "Gamma" }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item class used inside the data band.
public class Item
{
    public string Name { get; set; } = string.Empty;
}
