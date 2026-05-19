using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data with at least five items.
        var model = new ReportModel();
        model.Items.AddRange(new[]
        {
            new Item { Name = "Item 1" },
            new Item { Name = "Item 2" },
            new Item { Name = "Item 3" },
            new Item { Name = "Item 4" },
            new Item { Name = "Item 5" }, // This is the fifth item (index 4).
            new Item { Name = "Item 6" }
        });

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that uses ElementAt to fetch the fifth item.
        builder.Writeln("Fifth item: <<[model.Items.ElementAt(4).Name]>>");

        // Save the template (optional, demonstrates the load‑save cycle).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back (ensures BuildReport is called after loading).
        var loadedTemplate = new Document(templatePath);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}
