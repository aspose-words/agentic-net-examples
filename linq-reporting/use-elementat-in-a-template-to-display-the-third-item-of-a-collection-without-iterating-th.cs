using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the template document with LINQ Reporting tag.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Third item: <<[model.Items.ElementAt(2).Name]>>");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for reporting.
        // -------------------------------------------------
        var doc = new Document(templatePath);

        // -------------------------------------------------
        // Prepare sample data with at least three items.
        // -------------------------------------------------
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Item One" },
                new Item { Name = "Item Two" },
                new Item { Name = "Item Three" },
                new Item { Name = "Item Four" }
            }
        };

        // -------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // Save the generated report.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}
