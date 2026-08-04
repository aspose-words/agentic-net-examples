using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    // Initialize collection to avoid null reference.
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
                new Item { Name = "First" },
                new Item { Name = "Second" },
                new Item { Name = "Last" } // This should be displayed.
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that uses ElementAt to fetch the last item.
        // The expression calculates the index as Items.Count - 1.
        builder.Writeln("Last item name: <<[model.Items.ElementAt(model.Items.Count - 1).Name]>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var loadedDoc = new Document(templatePath);

        // Build the report using the model as the root data source named "model".
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        loadedDoc.Save(outputPath);
    }
}
