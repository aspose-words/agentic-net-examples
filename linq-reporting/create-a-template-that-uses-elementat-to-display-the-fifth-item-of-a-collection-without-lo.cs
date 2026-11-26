using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data with at least five items.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha" },
                new Item { Name = "Bravo" },
                new Item { Name = "Charlie" },
                new Item { Name = "Delta" },
                new Item { Name = "Echo" },      // Fifth item (index 4)
                new Item { Name = "Foxtrot" }
            }
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Fifth item: <<[model.Items.ElementAt(4).Name]>>");
        doc.Save(templatePath);

        // Load the template for report generation.
        var templateDoc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        var outputPath = "output.docx";
        templateDoc.Save(outputPath);
    }
}
