using System;
using System.Collections.Generic;
using System.IO;
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
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new() { Name = "Apple", Quantity = 5 },
                new() { Name = "Banana", Quantity = 0 },
                new() { Name = "Orange", Quantity = 3 },
                new() { Name = "Grapes", Quantity = -1 }
            }
        };

        // Create a template document with a LINQ Reporting tag that counts items with Quantity > 0.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Available items count (Quantity > 0): <<[Items.Count(i => i.Quantity > 0)]>>");
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
