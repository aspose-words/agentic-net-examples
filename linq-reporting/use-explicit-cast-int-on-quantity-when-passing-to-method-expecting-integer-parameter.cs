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
                new Item { Name = "Apple", Quantity = 5.7 },
                new Item { Name = "Banana", Quantity = 3.2 },
                new Item { Name = "Cherry", Quantity = 10.0 }
            }
        };

        // Create the template document and save it.
        const string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("report.docx");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Items Report");
        builder.Writeln();

        // Begin foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in model.Items]>>");

        // Output each item's name.
        builder.Writeln("Name: <<[item.Name]>>");

        // Output the quantity cast explicitly to int.
        builder.Writeln("Quantity (int): <<[(int)item.Quantity]>>");

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }

    // Root data model.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Item model used in the collection.
    public class Item
    {
        public string Name { get; set; } = "";
        public double Quantity { get; set; }
    }
}
