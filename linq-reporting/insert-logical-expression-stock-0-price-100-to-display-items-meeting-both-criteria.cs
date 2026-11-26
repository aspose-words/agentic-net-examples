using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Stock { get; set; }
    public decimal Price { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Items with Stock > 0 and Price < 100:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<if [item.Stock > 0 && item.Price < 100]>>");
        builder.Writeln("Name: <<[item.Name]>>, Stock: <<[item.Stock]>>, Price: <<[item.Price]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Name = "Apple", Stock = 10, Price = 0.99m },
                new Item { Name = "Banana", Stock = 0, Price = 0.59m },
                new Item { Name = "Cherry", Stock = 5, Price = 150.00m },
                new Item { Name = "Date", Stock = 3, Price = 75.50m }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
