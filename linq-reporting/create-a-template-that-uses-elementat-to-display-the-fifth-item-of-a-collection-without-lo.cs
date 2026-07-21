using System;
using System.Collections.Generic;
using System.IO;
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
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Item One" },
                new Item { Name = "Item Two" },
                new Item { Name = "Item Three" },
                new Item { Name = "Item Four" },
                new Item { Name = "Item Five" },
                new Item { Name = "Item Six" }
            }
        };

        // Define file names in the current working directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath   = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // Step 1: Create the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting tag that uses ElementAt to fetch the 5th item (index 4).
        builder.Writeln("Fifth item: <<[model.Items.ElementAt(4).Name]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model and the root name "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}
