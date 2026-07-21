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
        // 1. Create the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        // Insert a tag that counts items with Quantity > 0.
        builder.Writeln("Available items: <<[model.Items.Count(i => i.Quantity > 0)]>>");
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template for reporting.
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Apple",  Quantity = 5 },
                new Item { Name = "Banana", Quantity = 0 },
                new Item { Name = "Orange", Quantity = 3 },
                new Item { Name = "Pear",   Quantity = 0 }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);

        // Optional: output the count to the console for verification.
        int availableCount = model.Items.Count(i => i.Quantity > 0);
        Console.WriteLine($"Available items counted: {availableCount}");
    }
}
