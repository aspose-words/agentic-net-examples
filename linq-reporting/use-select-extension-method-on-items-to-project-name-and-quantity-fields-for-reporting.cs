using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Quantity { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a heading.
        builder.Writeln("Items Report");
        builder.Writeln();

        // Use the Select extension method to project only Name and Quantity.
        // The foreach tag iterates over the projected collection.
        builder.Writeln("<<foreach [item in Items.Select(i => new { i.Name, i.Quantity })]>>");
        builder.Writeln("Name: <<[item.Name]>>\tQuantity: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the data source.
        // -----------------------------------------------------------------
        List<Item> items = new()
        {
            new Item { Name = "Apple",  Quantity = 10 },
            new Item { Name = "Banana", Quantity = 20 },
            new Item { Name = "Orange", Quantity = 15 }
        };

        ReportModel model = new()
        {
            Items = items
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The root object name in the template is "model".
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        doc.Save(reportPath);
    }
}
