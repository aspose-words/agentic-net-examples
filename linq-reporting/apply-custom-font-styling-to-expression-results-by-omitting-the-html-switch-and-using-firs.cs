using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    // Name of the product.
    public string Name { get; set; } = string.Empty;

    // Quantity of the product.
    public int Quantity { get; set; }
}

public class ReportModel
{
    // Collection of items to be displayed in the report.
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln("==============");
        builder.Writeln();

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Apply custom font styling to the first character of the product name.
        // The first character is wrapped in a textColor tag (red), the rest is plain.
        // No -html switch is used; formatting is achieved via first‑character tags.
        builder.Writeln(
            "<<textColor [\"Red\"]>><<[item.Name.Substring(0,1)]>><</textColor>><<[item.Name.Substring(1)]>>" +
            " - Quantity: <<[item.Quantity]>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare the data source.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Quantity = 5 },
                new Item { Name = "Banana", Quantity = 3 },
                new Item { Name = "Cherry", Quantity = 7 }
            }
        };

        // -------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}
