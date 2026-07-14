using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Id { get; set; }
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
        var model = new ReportModel
        {
            Items =
            {
                new Item { Id = 1, Name = "Alpha" },
                new Item { Id = 2, Name = "Bravo" },
                new Item { Id = 3, Name = "Charlie" },
                new Item { Id = 4, Name = "Delta" },
                new Item { Id = 5, Name = "Echo" },   // Fifth item (index 4)
                new Item { Id = 6, Name = "Foxtrot" }
            }
        };

        // -----------------------------------------------------------------
        // Create a template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that uses ElementAt to fetch the 5th item.
        // ElementAt uses zero‑based indexing, so 4 corresponds to the fifth element.
        builder.Writeln("Fifth item: <<[model.Items.ElementAt(4).Name]>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back (ensures the document is fully prepared).
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
