using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();

        // Start a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Insert HTML that contains a centered paragraph for each item.
        builder.Writeln("<<html [item.Html]>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Text = "First paragraph" },
                new Item { Text = "Second paragraph" },
                new Item { Text = "Third paragraph" }
            }
        };

        // Populate the Html property with a centered paragraph for each item.
        foreach (var item in model.Items)
        {
            // Using simple HTML with inline style to achieve center alignment.
            item.Html = $"<p style=\"text-align:center;\">{System.Net.WebUtility.HtmlEncode(item.Text)}</p>";
        }

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item displayed in the report.
public class Item
{
    // Original text.
    public string Text { get; set; } = string.Empty;

    // HTML representation used in the template.
    public string Html { get; set; } = string.Empty;
}
