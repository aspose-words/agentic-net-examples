using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document with a LINQ Reporting tag that uses ElementAt to get the last item.
        var templatePath = "Template.docx";
        var outputPath = "Report.docx";

        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Last item: <<[Items.ElementAt(Items.Count - 1).Name]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha" },
                new Item { Name = "Beta" },
                new Item { Name = "Gamma" } // This should be displayed as the last item.
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
