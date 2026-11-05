using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data with at least five items
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new() { Name = "Item 1" },
                new() { Name = "Item 2" },
                new() { Name = "Item 3" },
                new() { Name = "Item 4" },
                new() { Name = "Item 5" },
                new() { Name = "Item 6" }
            }
        };

        // Create a template document programmatically
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Fifth item using ElementAt: <<[model.Items.ElementAt(4).Name]>>");
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
