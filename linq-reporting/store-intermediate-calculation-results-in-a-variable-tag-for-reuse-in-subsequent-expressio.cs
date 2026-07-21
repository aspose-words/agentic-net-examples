using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple",  Price = 1.20 },
                new Item { Name = "Banana", Price = 0.80 },
                new Item { Name = "Cherry", Price = 2.50 }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";

        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Iterate over the collection 'Items'.
        builder.Writeln("<<foreach [item in Items]>>");

        // Output each item's details.
        builder.Writeln("Item: <<[item.Name]>>  Price: <<[item.Price]>>");

        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Display the accumulated total using a LINQ expression.
        builder.Writeln("Total: <<[Items.Sum(p => p.Price)]>>");

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();

        // The root object name must match the name used in the template tags.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Data model classes (public with public properties, no nullable warnings)
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
